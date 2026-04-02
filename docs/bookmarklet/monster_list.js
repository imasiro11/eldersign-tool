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
    panel.textContent = "準備中...";
    document.body.appendChild(panel);
    return panel;
  };

  const parseOwnedCount = () => {
    const footer = document.querySelector("footer p");
    if (!footer) return null;
    const match = footer.textContent.match(/全\s*(\d+)\s*\/\s*\d+\s*枚/);
    if (!match) return null;
    const value = parseInt(match[1], 10);
    return Number.isNaN(value) ? null : value;
  };

  const getCurrentPage = () => {
    const on = document.querySelector("a.on");
    const value = on ? parseInt(on.textContent, 10) : 1;
    return Number.isNaN(value) ? 1 : value;
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

    const currentPage = getCurrentPage();
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

    const panel = buildPanel();
    const errors = [];
    const fragment = document.createDocumentFragment();

    const pagesToFetch = totalPages - currentPage;

    for (let page = currentPage + 1; page <= totalPages; page += 1) {
      panel.textContent = `取得中 ${page - currentPage}/${pagesToFetch} ページ`;
      try {
        const url = buildPageUrl(page);
        const response = await fetch(url, { credentials: "include" });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const html = await response.text();
        const items = extractListItems(html);
        if (!items.length) {
          throw new Error("一覧が取得できませんでした");
        }
        items.forEach((item) => fragment.appendChild(document.importNode(item, true)));
      } catch (err) {
        errors.push(`${page}ページ目: ${err.message}`);
      }
      await sleep(REQUEST_DELAY_MS);
    }

    list.appendChild(fragment);
    document.body.dataset.esBookExpanded = "1";

    if (errors.length) {
      panel.textContent = `完了: ${pagesToFetch}ページ / エラー: ${errors.length}`;
      console.warn("取得エラー", errors);
    } else {
      panel.textContent = "完了";
    }
    setTimeout(() => panel.remove(), 5000);
  };

  expandBookList().catch((err) => alert("エラー: " + err.message));
})();
