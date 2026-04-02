(() => {
  const PANEL_ID = "__es_book_expand_panel";
  const PER_PAGE = 10;
  const REQUEST_DELAY_MS = 500;

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const buildPanel = () => {
    document.getElementById(PANEL_ID)?.remove();
    const panel = document.createElement("div");
    panel.id = PANEL_ID;
    panel.style.cssText =
      "position:fixed;top:10px;right:10px;z-index:99999;" +
      "background:rgba(0,0,0,.8);color:#fff;padding:10px 12px;" +
      "border-radius:8px;font-family:monospace;font-size:12px;" +
      "max-width:calc(100% - 20px);";
    const status = document.createElement("div");
    status.textContent = "準備中...";
    const stopButton = document.createElement("button");
    stopButton.type = "button";
    stopButton.textContent = "停止";
    stopButton.style.cssText =
      "margin-top:8px;padding:4px 8px;border:0;border-radius:4px;" +
      "background:#c33;color:#fff;cursor:pointer;font:inherit;";
    panel.append(status, stopButton);
    document.body.appendChild(panel);
    return { panel, status, stopButton };
  };

  const parseOwnedCount = () => {
    const footer = document.querySelector("footer");
    const scope = footer || document.body;
    const texts = [...scope.querySelectorAll("p")].map((p) => p.textContent || "");
    for (const text of texts) {
      const normalized = text.replace(/\s+/g, "");
      const withTotal = normalized.match(/全\s*([\d,]+)\s*\/\s*([\d,]+)\s*(?:枚|人|件)?/);
      if (withTotal) {
        const value = parseInt(withTotal[1].replace(/,/g, ""), 10);
        return Number.isNaN(value) ? null : value;
      }
      const totalOnly = normalized.match(/全\s*([\d,]+)\s*(?:枚|人|件)?/);
      if (totalOnly) {
        const value = parseInt(totalOnly[1].replace(/,/g, ""), 10);
        return Number.isNaN(value) ? null : value;
      }
    }
    return null;
  };

  const findPagerContainer = () => {
    const footer = document.querySelector("footer");
    if (footer) {
      const pgBoxes = footer.querySelectorAll("div.pg");
      if (pgBoxes.length && pgBoxes[0].parentElement) {
        return pgBoxes[0].parentElement;
      }
      const nav = footer.querySelector("nav.pager");
      if (nav) return nav;
    }
    return document.querySelector("nav.pager");
  };

  const getCurrentPage = (pagerContainer) => {
    const currentLabel = pagerContainer?.querySelector("li.on a, a.on")?.textContent?.trim();
    if (currentLabel && /^\d+$/.test(currentLabel)) {
      const value = parseInt(currentLabel, 10);
      return Number.isNaN(value) ? 1 : value;
    }
    try {
      const url = new URL(location.href);
      const raw = url.searchParams.get("pg");
      const value = raw == null ? 1 : parseInt(raw, 10) + 1;
      return Number.isNaN(value) || value < 1 ? 1 : value;
    } catch (_) {
      return 1;
    }
  };

  const buildPageUrl = (pageNumber) => {
    const url = new URL(location.href);
    url.searchParams.set("pg", String(pageNumber - 1));
    return url.toString();
  };

  const extractListItems = (html) => {
    const doc = new DOMParser().parseFromString(html, "text/html");
    const items = doc.querySelectorAll("nav.block ul li");
    return Array.from(items);
  };

  const expandBookList = async () => {
    if (document.body.dataset.esBookExpanded === "1") {
      alert("すでに取得済みです。ページを再読み込みしてから実行してください。");
      return;
    }

    const list = document.querySelector("nav.block ul");
    if (!list) {
      alert("ブック一覧が見つかりません。ブック画面で実行してください。");
      return;
    }

    const pagerContainer = findPagerContainer();
    const currentPage = getCurrentPage(pagerContainer);
    const ownedCount = parseOwnedCount();
    if (!ownedCount) {
      alert("枚数の取得に失敗しました。");
      return;
    }

    const totalPages = Math.ceil(ownedCount / PER_PAGE);
    if (currentPage >= totalPages) {
      alert("追加ページはありません。");
      return;
    }

    const { panel, status, stopButton } = buildPanel();
    const errors = [];
    let isCanceled = false;
    let currentController = null;

    const pagesToFetch = totalPages - currentPage;
    stopButton.addEventListener("click", () => {
      isCanceled = true;
      currentController?.abort();
      status.textContent = "停止中...";
      stopButton.disabled = true;
      stopButton.style.cursor = "default";
    });

    for (let page = currentPage + 1; page <= totalPages; page += 1) {
      if (isCanceled) break;
      status.textContent = `取得中 ${page - currentPage}/${pagesToFetch} ページ`;
      try {
        const url = buildPageUrl(page);
        currentController = new AbortController();
        const response = await fetch(url, {
          credentials: "include",
          signal: currentController.signal,
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const html = await response.text();
        const items = extractListItems(html);
        if (!items.length) {
          throw new Error("一覧が取得できませんでした");
        }
        const fragment = document.createDocumentFragment();
        items.forEach((item) => fragment.appendChild(document.importNode(item, true)));
        list.appendChild(fragment);
      } catch (err) {
        if (isCanceled || err.name === "AbortError") {
          break;
        }
        errors.push(`${page}ページ目: ${err.message}`);
      } finally {
        currentController = null;
      }
      if (isCanceled) break;
      await sleep(REQUEST_DELAY_MS);
    }
    stopButton.remove();

    if (isCanceled) {
      status.textContent = "停止しました";
    } else {
      document.body.dataset.esBookExpanded = "1";
    }

    if (!isCanceled && errors.length) {
      status.textContent = `完了: ${pagesToFetch}ページ / エラー: ${errors.length}`;
      console.warn("取得エラー", errors);
    } else if (!isCanceled) {
      status.textContent = "完了";
    }
    setTimeout(() => panel.remove(), 5000);
  };

  expandBookList().catch((err) => alert("エラー: " + err.message));
})();
