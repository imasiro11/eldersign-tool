/**
 * 通常一覧・詳細絞り込み一覧のカードを選択し、保管庫／ブックへ移動する。
 * midリンクまたはitem_IDから個体を取得し、選択中は公式onclickによる遷移を止める。
 * 絞り込みで隠れた選択は解除する。移動成功時はDOMとcardDataを削除して公式一覧を再計算する。
 * 再実行・一括出品への切り替え・終了時はイベントと監視を解除する。
 */
(() => {
  const ACTION_PANEL_ID = "__es_book_action_panel";
  const STYLE_ID = "__es_book_select_style";
  const REQUEST_DELAY_MS = 500;
  const listeners = new AbortController();
  let selectionObserver;

  /** 絞り込みで非表示になった行は、一括選択・送信の対象に含めない。 */
  const isVisibleItem = (li) => !li.hidden && getComputedStyle(li).display !== "none";

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const PAGE_TYPES = {
    BOOK: "book",
    STORAGE: "storage",
    POST: "post",
  };

  const ensurePanelStyle = () => {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
#${ACTION_PANEL_ID}{
  position:fixed;
  bottom:40px;
  left:50%;
  transform:translateX(-50%);
  z-index:99999;
  background:rgba(0,0,0,.8);
  color:#fff;
  padding:8px 10px;
  border-radius:8px;
  font-family:monospace;
  font-size:12px;
  width:calc(100% - 12px);
  max-width:420px;
  display:flex;
  flex-direction:column;
  align-items:stretch;
  box-sizing:border-box;
}
@media (max-width: 767px){
  #${ACTION_PANEL_ID}{
    top:72px;
    bottom:auto;
  }
}
#${ACTION_PANEL_ID} button,
#${ACTION_PANEL_ID} input{
  font:inherit;
}
#${ACTION_PANEL_ID} .es-menu-close{
  position:absolute;
  top:0;
  right:0;
  cursor:pointer;
  color:#fff;
  font-size:16px;
  line-height:16px;
  padding:6px 10px;
  border:0;
  background:transparent;
}
#${ACTION_PANEL_ID} .es-book-menu-list{
  display:flex;
  gap:6px;
  flex-wrap:wrap;
  align-items:center;
  align-content:flex-start;
  overflow-y:auto;
  max-height:320px;
  padding-top:8px;
  padding-right:2px;
  width:100%;
  box-sizing:border-box;
}
#${ACTION_PANEL_ID} .es-book-status{
  display:inline-flex;
  align-items:center;
  line-height:1;
  min-height:31px;
}
.es-book-item{position:relative;}
.es-book-item a{cursor:pointer;display:block;border-radius:12px;position:relative;padding-right:40px;}
.es-book-check{position:absolute;right:14px;top:50%;transform:translateY(-50%);}
.es-book-item.is-selected a{
  box-shadow:0 0 0 2px #f3e6c1 inset,0 0 0 4px rgba(81,51,18,.8);
}
`.trim();
    document.head.appendChild(style);
  };

  const buildActionPanel = () => {
    ensurePanelStyle();
    document.getElementById(ACTION_PANEL_ID)?.remove();
    const panel = document.createElement("div");
    panel.id = ACTION_PANEL_ID;

    const makeButton = (label) => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = label;
      return button;
    };

    const storeButton = makeButton("保管する");
    const bookButton = makeButton("ブックへ");
    const selectAllButton = makeButton("全選択");
    const selectUnprotectedButton = makeButton("保護以外選択");
    const closeButton = document.createElement("span");
    closeButton.className = "es-menu-close";
    closeButton.setAttribute("role", "button");
    closeButton.setAttribute("tabindex", "0");
    closeButton.title = "閉じる";
    closeButton.innerHTML =
      '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="currentColor">' +
      '<path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>';
    const status = document.createElement("span");
    status.className = "es-book-status";
    status.textContent = "選択待ち";
    closeButton.setAttribute("aria-label", "閉じる");
    const menuList = document.createElement("div");
    menuList.className = "es-book-menu-list";

    menuList.appendChild(selectAllButton);
    menuList.appendChild(selectUnprotectedButton);
    menuList.appendChild(bookButton);
    menuList.appendChild(storeButton);
    menuList.appendChild(status);
    panel.appendChild(closeButton);
    panel.appendChild(menuList);
    document.body.appendChild(panel);

    return {
      panel,
      storeButton,
      bookButton,
      selectAllButton,
      selectUnprotectedButton,
      closeButton,
      menuList,
      status,
    };
  };

  const parseMidFromHref = (href) => {
    if (!href) return null;
    try {
      const url = new URL(href, location.href);
      const mid = url.searchParams.get("mid");
      return mid && /^\d+$/.test(mid) ? mid : null;
    } catch (err) {
      return null;
    }
  };

  const updateSelectedCount = (status) => {
    const count = collectSelected().length;
    status.textContent = `選択中: ${count}枚`;
  };

  const toggleSelection = (li, checkbox, status) => {
    checkbox.checked = !checkbox.checked;
    li.classList.toggle("is-selected", checkbox.checked);
    updateSelectedCount(status);
  };

  const updateSelectionState = (li, checkbox, status) => {
    li.classList.toggle("is-selected", checkbox.checked);
    updateSelectedCount(status);
  };

  /** 通常リンクのmidと詳細一覧のitem_IDを扱い、カード全体のタップを選択操作へ切り替える。 */
  const setupSelectableList = (list, status, pageType) => {
    ensurePanelStyle();
    list.querySelectorAll("li").forEach((li) => {
      if (li.classList.contains("es-book-item")) return;
      const anchor = li.querySelector("a");
      if (!anchor) return;
      const mid = parseMidFromHref(anchor.href) || li.id.match(/^item_(\d+)$/)?.[1];
      if (!mid) return;
      if (pageType === PAGE_TYPES.POST && !anchor.href.includes("stockbox?cmd=gmc")) return;

      li.classList.add("es-book-item");
      li.dataset.mid = mid;
      li.dataset.href = anchor.href;

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "es-book-check";
      li.appendChild(checkbox);
      checkbox.addEventListener("change", () => {
        updateSelectionState(li, checkbox, status);
      }, { signal: listeners.signal });
    });

    if (!list.dataset.esBookSelectable) {
      list.dataset.esBookSelectable = "1";
      list.addEventListener("click", (event) => {
        const rawTarget = event.target;
        const target =
          rawTarget instanceof Element
            ? rawTarget
            : rawTarget instanceof Node
              ? rawTarget.parentElement
              : null;
        if (!target) return;
        const li = target.closest("li.es-book-item");
        if (!li || !list.contains(li)) return;
        if (target.closest("button")) return;
        const checkbox = li.querySelector("input.es-book-check");
        if (!checkbox) return;
        event.stopImmediatePropagation();
        if (target === checkbox) return;
        event.preventDefault();
        toggleSelection(li, checkbox, status);
      }, { capture: true, signal: listeners.signal });
    }
  };

  const collectSelected = () => {
    return Array.from(document.querySelectorAll("li.es-book-item.is-selected"))
      .filter(isVisibleItem)
      .map((li) => {
        const mid = li.dataset.mid;
        if (!mid) return null;
        return { mid, href: li.dataset.href || "", li };
      })
      .filter(Boolean);
  };

  const teardownSelectableList = (list) => {
    listeners.abort();
    selectionObserver?.disconnect();
    if (window.__eldersignCardSelectionCleanup === cleanup) delete window.__eldersignCardSelectionCleanup;
    document.getElementById(ACTION_PANEL_ID)?.remove();
    document.getElementById(STYLE_ID)?.remove();
    list.querySelectorAll("li.es-book-item").forEach((li) => {
      li.classList.remove("es-book-item", "is-selected");
      delete li.dataset.mid;
      delete li.dataset.href;
      li.querySelector("input.es-book-check")?.remove();
    });
    delete list.dataset.esBookSelectable;
  };

  /** 詳細一覧には非表示の状態アイコンも存在するため、表示中のアイコンだけで状態を判定する。 */
  const getItemFlags = (li) => {
    const icons = Array.from(li.querySelectorAll("img.i")).filter((icon) => getComputedStyle(icon).display !== "none");
    if (!icons.length) return { isOnSale: false, isProtected: false, isForming: false };
    const srcList = icons.map((icon) => icon.getAttribute("src") || "");
    return {
      isOnSale: srcList.some((src) => src.includes("card_b")),
      isProtected: srcList.some((src) => src.includes("card_l")),
      isForming: srcList.some((src) => src.includes("card_c")),
    };
  };

  const setAllSelections = (checked, status, options = {}) => {
    const { skipOnSale = false, skipProtected = false, skipForming = false } = options;
    const items = document.querySelectorAll("li.es-book-item");
    items.forEach((li) => {
      if (checked && !isVisibleItem(li)) return;
      if (checked && (skipOnSale || skipProtected || skipForming)) {
        const { isOnSale, isProtected, isForming } = getItemFlags(li);
        if (
          (skipOnSale && isOnSale) ||
          (skipProtected && isProtected) ||
          (skipForming && isForming)
        ) return;
      }
      const checkbox = li.querySelector("input.es-book-check");
      if (!checkbox) return;
      checkbox.checked = checked;
      updateSelectionState(li, checkbox, status);
    });
  };

  const buildMoveUrl = (action, selection, pageType) => {
    const { mid, href } = selection;

    if (pageType === PAGE_TYPES.POST) {
      if (action !== "book") {
        throw new Error("ポストから保管庫への移動は未対応です。");
      }
      const sourceUrl = new URL(href, location.href);
      sourceUrl.searchParams.set("cmd", "gm");
      sourceUrl.searchParams.set("mid", mid);
      sourceUrl.searchParams.delete("srt");
      return sourceUrl.toString();
    }

    const url = new URL("/mcard_detail", location.origin);
    url.searchParams.set("cmd", action === "storage" ? "a1" : "a2");
    url.searchParams.set("mid", mid);
    if (action === "book") {
      url.searchParams.set("ex", "1");
    }
    return url.toString();
  };

  const runMove = async (action, status, pageType, onComplete) => {
    const selections = collectSelected();
    if (!selections.length) {
      alert("選択されたモンスターがありません。");
      return;
    }

    let errors = 0;
    let shouldStop = false;
    let lastErrorMessage = "";
    for (let i = 0; i < selections.length; i += 1) {
      if (shouldStop) break;
      status.textContent = `送信中 ${i + 1}/${selections.length}`;
      const selection = selections[i];
      const { mid, li } = selection;
      try {
        const url = buildMoveUrl(action, selection, pageType);
        const response = await fetch(url, { credentials: "include" });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const html = await response.text();
        const doc = new DOMParser().parseFromString(html, "text/html");
        const errArticle = doc.querySelector("article.err");
        if (errArticle) {
          errors += 1;
          const message =
            errArticle.querySelector("p")?.textContent?.trim() ||
            "移動できませんでした。";
          status.textContent = message;
          lastErrorMessage = message;
          if (message.includes("一杯です")) {
            shouldStop = true;
          }
        } else {
          // 公式ソートが移動済みカードのDOMを参照しないよう、一覧データも同期する。
          if (listIsFilterScreen() && Array.isArray(window.cardData)) {
            const index = window.cardData.findIndex((card) => String(card.uid) === mid);
            if (index >= 0) window.cardData.splice(index, 1);
          }
          li.remove();
          if (listIsFilterScreen()) window.mcS?.SortFunc?.ApplySort();
          updateSelectedCount(status);
        }
      } catch (err) {
        errors += 1;
        console.warn("移動エラー", mid, err);
      }
      await sleep(REQUEST_DELAY_MS);
    }

    if (errors && lastErrorMessage) {
      status.textContent = lastErrorMessage;
    } else {
      status.textContent = errors
        ? `完了(エラー:${errors})`
        : "完了";
    }
    if (errors === 0) {
      onComplete?.();
    }
  };

  /** 公式の詳細一覧データを持つ画面か判定する。 */
  const listIsFilterScreen = () => !!document.getElementById("mlist");
  let cleanup;

  const init = async () => {
    const list = document.querySelector("#mlist, nav.block ul");
    if (!list) return;
    window.__eldersignCardSelectionCleanup?.();
    cleanup = () => teardownSelectableList(list);
    window.__eldersignCardSelectionCleanup = cleanup;
    const {
      storeButton,
      bookButton,
      selectAllButton,
      selectUnprotectedButton,
      closeButton,
      status,
    } =
      buildActionPanel();
    const heading = document.querySelector("h1");
    const headingText = heading ? heading.textContent.trim() : "";
    const pageType =
      headingText === "保管庫"
        ? PAGE_TYPES.STORAGE
        : headingText === "ブック"
          ? PAGE_TYPES.BOOK
          : headingText === "ポスト"
            ? PAGE_TYPES.POST
            : null;
    if (!pageType) return;
    if (pageType === PAGE_TYPES.STORAGE || pageType === PAGE_TYPES.POST) {
      storeButton.style.display = "none";
    } else if (pageType === PAGE_TYPES.BOOK) {
      bookButton.style.display = "none";
    }
    if (pageType === PAGE_TYPES.POST) {
      selectUnprotectedButton.style.display = "none";
    }
    let nextSelectAll = true;
    setupSelectableList(list, status, pageType);
    selectionObserver = new MutationObserver(() => {
      let changed = false;
      list.querySelectorAll("li.es-book-item.is-selected").forEach((li) => {
        if (!isVisibleItem(li)) {
          li.classList.remove("is-selected");
          li.querySelector(".es-book-check").checked = false;
          changed = true;
        }
      });
      if (changed) updateSelectedCount(status);
    });
    selectionObserver.observe(list, { subtree: true, attributes: true, attributeFilter: ["style", "hidden"], childList: true });
    updateSelectedCount(status);
    storeButton.addEventListener("click", () =>
      runMove("storage", status, pageType, () => teardownSelectableList(list))
    );
    bookButton.addEventListener("click", () =>
      runMove("book", status, pageType, () => teardownSelectableList(list))
    );
    selectAllButton.addEventListener("click", () => {
      if (nextSelectAll) {
        setAllSelections(true, status, { skipOnSale: true, skipForming: true });
        selectAllButton.textContent = "全削除";
      } else {
        setAllSelections(false, status);
        selectAllButton.textContent = "全選択";
      }
      nextSelectAll = !nextSelectAll;
    });
    selectUnprotectedButton.addEventListener("click", () => {
      setAllSelections(true, status, {
        skipOnSale: true,
        skipProtected: true,
        skipForming: true,
      });
    });
    closeButton.addEventListener("click", () => {
      teardownSelectableList(list);
    });
    closeButton.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        teardownSelectableList(list);
      }
    });
  };

  init().catch((err) => alert("エラー: " + err.message));
})();
