(() => {
  const ACTION_PANEL_ID = "__es_bazaar_listing_panel";
  const STYLE_ID = "__es_bazaar_listing_style";
  const REQUEST_DELAY_MS = 500;
  const MAX_PRICE_ANY = 100000000;
  const MULTIPLIER_STORAGE_KEY = "esBazaarBulkListingMultiplier";
  const TARGET_STATS = ["HP", "攻撃", "魔力", "防御", "命中", "敏捷"];
  const RARITY_COEFFICIENT_BY_MAX_LEVEL = Object.freeze({
    30: 1,
    50: 1.5,
    70: 2,
    90: 2.5,
  });

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const loadSavedMultiplier = () => {
    try {
      const savedValue = localStorage.getItem(MULTIPLIER_STORAGE_KEY);
      if (savedValue === null || savedValue.trim() === "") return "";
      const multiplier = Number(savedValue);
      return Number.isFinite(multiplier) && multiplier >= 0 ? savedValue : "";
    } catch (error) {
      console.warn("保存した倍率を読み込めませんでした。", error);
      return "";
    }
  };

  const saveMultiplier = (value) => {
    try {
      localStorage.setItem(MULTIPLIER_STORAGE_KEY, value);
    } catch (error) {
      console.warn("倍率を保存できませんでした。", error);
    }
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
@media (max-width:767px){
  #${ACTION_PANEL_ID}{top:72px;bottom:auto;}
}
#${ACTION_PANEL_ID} button,
#${ACTION_PANEL_ID} input{font:inherit;}
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
#${ACTION_PANEL_ID} .es-bazaar-menu-list{
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
#${ACTION_PANEL_ID} .es-bazaar-multiplier{
  display:inline-flex;
  align-items:center;
  gap:4px;
}
#${ACTION_PANEL_ID} .es-bazaar-multiplier input{width:64px;}
#${ACTION_PANEL_ID} .es-bazaar-row-break{
  flex-basis:100%;
  height:0;
}
#${ACTION_PANEL_ID} .es-bazaar-rename{
  display:inline-flex;
  align-items:center;
  gap:4px;
}
#${ACTION_PANEL_ID} .es-bazaar-status{
  display:inline-flex;
  align-items:center;
  line-height:1.35;
  min-height:31px;
}
.es-bazaar-item{position:relative;}
.es-bazaar-item a{cursor:pointer;display:block;border-radius:12px;position:relative;padding-right:40px;}
.es-bazaar-check{position:absolute;right:14px;top:50%;transform:translateY(-50%);}
.es-bazaar-item.is-selected a{
  box-shadow:0 0 0 2px #f3e6c1 inset,0 0 0 4px rgba(81,51,18,.8);
}
`.trim();
    document.head.appendChild(style);
  };

  const makeButton = (label) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = label;
    return button;
  };

  const buildActionPanel = () => {
    ensurePanelStyle();
    document.getElementById(ACTION_PANEL_ID)?.remove();

    const panel = document.createElement("div");
    panel.id = ACTION_PANEL_ID;
    const closeButton = document.createElement("span");
    closeButton.className = "es-menu-close";
    closeButton.setAttribute("role", "button");
    closeButton.setAttribute("tabindex", "0");
    closeButton.setAttribute("aria-label", "閉じる");
    closeButton.title = "閉じる";
    closeButton.innerHTML =
      '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" fill="currentColor">' +
      '<path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>';

    const menuList = document.createElement("div");
    menuList.className = "es-bazaar-menu-list";
    const multiplierLabel = document.createElement("label");
    multiplierLabel.className = "es-bazaar-multiplier";
    multiplierLabel.append("倍率");
    const multiplierInput = document.createElement("input");
    multiplierInput.type = "number";
    multiplierInput.min = "0";
    multiplierInput.step = "0.01";
    multiplierInput.placeholder = "例: 10";
    multiplierInput.value = loadSavedMultiplier();
    multiplierInput.inputMode = "decimal";
    multiplierLabel.appendChild(multiplierInput);

    const renameLabel = document.createElement("label");
    renameLabel.className = "es-bazaar-rename";
    const renameInput = document.createElement("input");
    renameInput.type = "checkbox";
    renameInput.checked = false;
    renameLabel.append(renameInput, "呼称を納品基本ptへ変更");
    const rowBreak = document.createElement("span");
    rowBreak.className = "es-bazaar-row-break";
    rowBreak.setAttribute("aria-hidden", "true");

    const selectAllButton = makeButton("全選択");
    const listButton = makeButton("出品する");
    const status = document.createElement("span");
    status.className = "es-bazaar-status";
    status.textContent = "選択中: 0枚";

    menuList.append(multiplierLabel, selectAllButton, listButton, status, rowBreak, renameLabel);
    panel.append(closeButton, menuList);
    document.body.appendChild(panel);
    return {
      panel,
      closeButton,
      multiplierInput,
      renameInput,
      selectAllButton,
      listButton,
      status,
    };
  };

  const parseMonsterId = (href) => {
    if (!href) return null;
    try {
      const id = new URL(href, location.href).searchParams.get("mid");
      return id && /^\d+$/.test(id) ? id : null;
    } catch (_) {
      return null;
    }
  };

  const getItemFlags = (li) => {
    const sources = [...li.querySelectorAll("img.i")].map((image) => image.getAttribute("src") || "");
    return {
      isOnSale: sources.some((src) => src.includes("card_b")),
      isProtected: sources.some((src) => src.includes("card_l")),
      isForming: sources.some((src) => src.includes("card_c")),
    };
  };

  const isUnavailable = (li) => {
    const flags = getItemFlags(li);
    return flags.isProtected || flags.isForming;
  };

  const isAlreadyListed = (li) => {
    const recordedState = li.dataset.esBazaarListingState;
    if (recordedState) return recordedState === "listed";
    return getItemFlags(li).isOnSale;
  };

  const updateSelectedCount = (status) => {
    const count = document.querySelectorAll("li.es-bazaar-item.is-selected").length;
    status.textContent = `選択中: ${count}枚`;
  };

  const updateSelection = (li, checkbox, status) => {
    li.classList.toggle("is-selected", checkbox.checked);
    updateSelectedCount(status);
  };

  const setupSelectableList = (list, status) => {
    list.querySelectorAll("li").forEach((li) => {
      const anchor = li.querySelector("a");
      const monsterId = parseMonsterId(anchor?.href);
      if (!anchor || !monsterId) return;

      li.classList.add("es-bazaar-item");
      li.dataset.monsterId = monsterId;
      li.dataset.detailUrl = anchor.href;
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "es-bazaar-check";
      checkbox.disabled = isUnavailable(li);
      checkbox.setAttribute("aria-label", `${li.querySelector("h1")?.textContent?.trim() || monsterId}を選択`);
      anchor.appendChild(checkbox);
      checkbox.addEventListener("change", () => updateSelection(li, checkbox, status));
    });

    list.addEventListener("click", (event) => {
      const target = event.target instanceof Element ? event.target : event.target?.parentElement;
      const li = target?.closest("li.es-bazaar-item");
      if (!li || !list.contains(li) || target === li.querySelector("input.es-bazaar-check")) return;
      event.preventDefault();
      const checkbox = li.querySelector("input.es-bazaar-check");
      if (checkbox.disabled) return;
      checkbox.checked = !checkbox.checked;
      updateSelection(li, checkbox, status);
    });
  };

  const collectSelected = () =>
    [...document.querySelectorAll("li.es-bazaar-item.is-selected")].map((li) => ({
      li,
      monsterId: li.dataset.monsterId,
      detailUrl: li.dataset.detailUrl,
      name: li.querySelector("h1")?.textContent?.trim() || `ID:${li.dataset.monsterId}`,
    }));

  const setAllSelections = (checked, status) => {
    document.querySelectorAll("li.es-bazaar-item").forEach((li) => {
      const checkbox = li.querySelector("input.es-bazaar-check");
      if (!checkbox) return;
      checkbox.checked = checked && !isUnavailable(li);
      updateSelection(li, checkbox, status);
    });
  };

  const teardown = (list) => {
    document.getElementById(ACTION_PANEL_ID)?.remove();
    document.getElementById(STYLE_ID)?.remove();
    list.querySelectorAll("li.es-bazaar-item").forEach((li) => {
      li.classList.remove("es-bazaar-item", "is-selected");
      delete li.dataset.monsterId;
      delete li.dataset.detailUrl;
      li.querySelector("input.es-bazaar-check")?.remove();
    });
  };

  const getLevelInfo = (doc) => {
    const heading = doc.querySelector("div.card_d header.card h3");
    const match = heading?.textContent.match(/Lv\s*(\d+)\s*\/\s*(\d+)/i);
    if (!match) throw new Error("レベル情報を取得できませんでした。");
    return { level: Number(match[1]), maxLevel: Number(match[2]) };
  };

  const getSeraenoPromotion = (doc) => {
    const src = doc.querySelector("img#card")?.getAttribute("src") || "";
    const match = src.split(/[?#]/, 1)[0].match(/_[0-3]g\d+([gs])\.[a-z0-9]+$/i);
    if (!match) return null;
    return match[1].toLowerCase() === "g"
      ? { sourceCoefficient: 1.5, sourceMaxLevel: 50 }
      : { sourceCoefficient: 1, sourceMaxLevel: 30 };
  };

  const getGrowth = (level, maxLevel, rarityCoefficient, promotion) => {
    let growth;
    if (promotion && promotion.sourceMaxLevel > 1 && promotion.sourceMaxLevel < maxLevel) {
      if (level <= promotion.sourceMaxLevel) {
        const sourceProgress = Math.max(0, (level - 1) / (promotion.sourceMaxLevel - 1));
        growth = 1 + promotion.sourceCoefficient * sourceProgress;
      } else {
        const promotedProgress =
          (level - promotion.sourceMaxLevel) / (maxLevel - promotion.sourceMaxLevel);
        growth =
          1 +
          promotion.sourceCoefficient +
          (rarityCoefficient - promotion.sourceCoefficient) * promotedProgress;
      }
    } else {
      growth = 1 + rarityCoefficient * Math.max(0, (level - 1) / (maxLevel - 1));
    }
    return { growth, hpGrowth: Math.sqrt(growth) };
  };

  const estimateLevelOneBase = (currentTotal, bonus, isHp, level, growth, hpGrowth) => {
    if (currentTotal <= 0) return 0;
    if (level <= 1) return Math.max(1, currentTotal - bonus);
    const factor = isHp ? hpGrowth : growth;
    const approximate = currentTotal / factor - bonus;
    let best = Math.max(1, Math.floor(approximate));
    let bestDifference = Math.abs(Math.floor((best + bonus) * factor) - currentTotal);
    let maximumExact = null;
    for (let candidate = Math.max(1, Math.floor(approximate) - 50); candidate <= Math.floor(approximate) + 50; candidate += 1) {
      const simulated = Math.floor((candidate + bonus) * factor);
      const difference = Math.abs(simulated - currentTotal);
      if (simulated === currentTotal) {
        maximumExact = maximumExact === null ? candidate : Math.max(maximumExact, candidate);
      } else if (difference < bestDifference) {
        bestDifference = difference;
        best = candidate;
      }
    }
    return maximumExact ?? best;
  };

  const calculateDeliveryPoint = (doc) => {
    const { level, maxLevel } = getLevelInfo(doc);
    const rarityCoefficient = RARITY_COEFFICIENT_BY_MAX_LEVEL[maxLevel] ?? 1;
    const { growth, hpGrowth } = getGrowth(
      level,
      maxLevel,
      rarityCoefficient,
      getSeraenoPromotion(doc)
    );
    const caption = [...doc.querySelectorAll("div.status table caption")].find(
      (element) => element.textContent.trim() === "ステータス"
    );
    if (!caption) throw new Error("ステータスを取得できませんでした。");

    let sumOfSquares = 0;
    let statCount = 0;
    caption.closest("table").querySelectorAll("tbody tr").forEach((row) => {
      const label = row.querySelector("th")?.textContent.trim();
      if (!TARGET_STATS.includes(label)) return;
      const cells = row.querySelectorAll("td");
      if (cells.length < 2) return;
      const parts = cells[0].textContent.trim().split("/");
      const valueText = label === "HP" && parts.length >= 2 ? parts[1] : parts[0];
      const currentTotal = Number.parseInt(valueText.replace(/[^\d-]/g, ""), 10);
      const bonusMatch = cells[1].textContent.match(/([+-]?\d+)/);
      const bonus = bonusMatch ? Number.parseInt(bonusMatch[1], 10) : 0;
      if (!Number.isFinite(currentTotal)) return;
      const base = estimateLevelOneBase(
        currentTotal,
        bonus,
        label === "HP",
        level,
        growth,
        hpGrowth
      );
      if (base <= 0) return;
      sumOfSquares += (bonus / base) ** 2;
      statCount += 1;
    });
    if (statCount !== TARGET_STATS.length) throw new Error("6項目のステータスを取得できませんでした。");

    const rawEvaluation = Math.sqrt(sumOfSquares / 6) * 200 + 10;
    const evaluation = Math.floor(rawEvaluation * 10) / 10;
    return Math.floor(evaluation * rarityCoefficient);
  };

  const fetchDocument = async (url) => {
    const response = await fetch(url, { credentials: "include" });
    if (!response.ok) throw new Error(`通信に失敗しました（HTTP ${response.status}）。`);
    const doc = new DOMParser().parseFromString(await response.text(), "text/html");
    const errorArticle = doc.querySelector("article.err");
    if (errorArticle) throw new Error(errorArticle.textContent.trim() || "操作を完了できませんでした。");
    return doc;
  };

  const findListingPageUrl = (detailDoc, selection) => {
    const link = detailDoc.querySelector('a[href*="mcard_sell"]');
    if (link) return new URL(link.getAttribute("href"), location.origin).toString();

    if (!selection.monsterId) {
      throw new Error("出品画面のURLを作成できませんでした。");
    }
    const listingUrl = new URL("/mcard_sell", location.origin);
    listingUrl.searchParams.set("mid", selection.monsterId);
    try {
      const page = new URL(selection.detailUrl, location.href).searchParams.get("pg");
      if (page !== null) listingUrl.searchParams.set("pg", page);
    } catch (_) {
      // ページ番号が取得できなくても出品画面は開けるため、そのまま続行する。
    }
    return listingUrl.toString();
  };

  const buildSubmitUrl = (listingDoc, price) => {
    const scripts = [...listingDoc.querySelectorAll("script")]
      .map((script) => script.textContent || "")
      .join("\n");
    const match = scripts.match(/["']([^"']*mcard_detail\?[^"']*\bcmd=s[^"']*\bany=)["']/i);
    if (!match) throw new Error("出品用の送信先を取得できませんでした。");
    const url = new URL(match[1], location.origin);
    url.searchParams.set("any", String(price));
    return url.toString();
  };

  const getListedPrice = (doc) => {
    const emphasizedPrice = doc.querySelector("article.checktxt em")?.textContent || "";
    const emphasizedMatch = emphasizedPrice.match(/([\d,]+)\s*Any/i);
    if (emphasizedMatch) return Number.parseInt(emphasizedMatch[1].replace(/,/g, ""), 10);

    const bodyText = doc.body?.textContent || "";
    const textMatch = bodyText.match(/(?:出品金額は|バザーに)\s*([\d,]+)\s*Any/i);
    if (textMatch) return Number.parseInt(textMatch[1].replace(/,/g, ""), 10);

    const scripts = [...doc.querySelectorAll("script")]
      .map((script) => script.textContent || "")
      .join("\n");
    const scriptMatch = scripts.match(/confirm\.set\([^)]*?["']([\d,]+)\s*Anyで購入します/i);
    return scriptMatch ? Number.parseInt(scriptMatch[1].replace(/,/g, ""), 10) : null;
  };

  const getCurrentPrice = (listingDoc) => {
    const text = listingDoc.querySelector("section.chr")?.textContent || "";
    const match = text.match(/現在の価格:\s*([\d,]+)\s*Any/i);
    return match ? Number.parseInt(match[1].replace(/,/g, ""), 10) : null;
  };

  const renameMonster = async (monsterId, nextName) => {
    const renameUrl = new URL("/mcard_name", location.origin);
    renameUrl.searchParams.set("mid", monsterId);
    renameUrl.searchParams.set("pg", "0");
    const renamePage = await fetchDocument(renameUrl.toString());
    const form = renamePage.querySelector('form[action*="mcard_detail"]');
    const action = form?.getAttribute("action");
    if (!action) throw new Error("呼称変更フォームを取得できませんでした。");

    const response = await fetch(new URL(action, location.origin).toString(), {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ name: nextName }).toString(),
    });
    if (!response.ok) throw new Error(`呼称変更に失敗しました（HTTP ${response.status}）。`);
    const resultDoc = new DOMParser().parseFromString(await response.text(), "text/html");
    const errorArticle = resultDoc.querySelector("article.err");
    if (errorArticle) throw new Error(errorArticle.textContent.trim() || "呼称変更に失敗しました。");
    return nextName;
  };

  const appendListingResult = (li, price, wasAlreadyUnlisted) => {
    let priceElement = li.querySelector(".es-bazaar-listed-price");
    if (!priceElement) {
      priceElement = document.createElement("p");
      priceElement.className = "es-bazaar-listed-price";
      const clearElement = li.querySelector("a > div[style*='clear']");
      if (clearElement) clearElement.before(priceElement);
      else li.querySelector("a")?.appendChild(priceElement);
    }
    priceElement.textContent =
      price === 0
        ? wasAlreadyUnlisted
          ? "出品なし"
          : "出品取り下げ済み"
        : `出品価格: ${price.toLocaleString()} Any`;
  };

  const listOne = async (selection, multiplier, shouldRename) => {
    if (isUnavailable(selection.li)) {
      throw new Error("保護中・編成中のカードです。");
    }
    const detailDoc = await fetchDocument(selection.detailUrl);
    const deliveryPoint = calculateDeliveryPoint(detailDoc);
    const detailListedPrice = getListedPrice(detailDoc);
    const alreadyListed =
      detailListedPrice !== null ? detailListedPrice > 0 : isAlreadyListed(selection.li);
    const price =
      multiplier === 0
        ? 0
        : Math.min(MAX_PRICE_ANY, Math.max(50, Math.round(deliveryPoint * multiplier)));
    if (shouldRename && !alreadyListed) {
      const nextName = await renameMonster(selection.monsterId, `納品基本pt${deliveryPoint}`);
      const heading = selection.li.querySelector("h1");
      if (heading) heading.textContent = nextName;
      selection.name = nextName;
    }
    const listingDoc = await fetchDocument(findListingPageUrl(detailDoc, selection));
    if (price === 0) {
      const currentPrice = getCurrentPrice(listingDoc);
      if (currentPrice === null) throw new Error("現在の出品価格を取得できませんでした。");
      if (currentPrice === 0) {
        return { deliveryPoint, price: 0, wasAlreadyUnlisted: true };
      }
    }
    const resultDoc = await fetchDocument(buildSubmitUrl(listingDoc, price));
    if (resultDoc.querySelector("#any") || !resultDoc.querySelector("div.card_d")) {
      throw new Error("出品が完了したことを確認できませんでした。");
    }
    return {
      deliveryPoint,
      price: getListedPrice(resultDoc) ?? price,
      wasAlreadyUnlisted: false,
    };
  };

  const runListing = async (controls) => {
    const selections = collectSelected();
    if (!selections.length) {
      alert("選択されたモンスターがありません。");
      return;
    }
    const multiplierText = controls.multiplierInput.value.trim();
    const multiplier = Number(multiplierText);
    if (multiplierText === "" || !Number.isFinite(multiplier) || multiplier < 0) {
      alert("倍率には0以上の数値を入力してください。");
      controls.multiplierInput.focus();
      return;
    }
    const shouldRename = controls.renameInput.checked;
    const renameMessage = shouldRename ? " 呼称も納品基本ptへ変更します。" : "";
    const actionMessage =
      multiplier === 0
        ? `${selections.length}枚の出品を取り下げます。未出品のカードは変更しません。`
        : `${selections.length}枚を納品基本pt × ${multiplier} Anyで出品します。`;
    if (!confirm(`${actionMessage}${renameMessage}よろしいですか？`)) return;
    saveMultiplier(multiplierText);

    controls.listButton.disabled = true;
    controls.selectAllButton.disabled = true;
    controls.multiplierInput.disabled = true;
    controls.renameInput.disabled = true;
    controls.closeButton.style.display = "none";
    try {
      let errors = 0;
      for (let index = 0; index < selections.length; index += 1) {
        const selection = selections[index];
        controls.status.textContent = `処理中 ${index + 1}/${selections.length}: ${selection.name}`;
        try {
          const result = await listOne(selection, multiplier, shouldRename);
          selection.li.classList.remove("is-selected");
          const checkbox = selection.li.querySelector("input.es-bazaar-check");
          if (checkbox) {
            checkbox.checked = false;
          }
          selection.li.dataset.esBazaarListingState =
            result.price === 0 ? "unlisted" : "listed";
          appendListingResult(selection.li, result.price, result.wasAlreadyUnlisted);
          controls.status.textContent =
            result.price === 0
              ? result.wasAlreadyUnlisted
                ? `${selection.name}: 出品されていません`
                : `${selection.name}: 出品を取り下げました`
              : `${selection.name}: ${result.deliveryPoint}pt × ${multiplier} = ${result.price} Any`;
        } catch (error) {
          errors += 1;
          console.warn("一括バザー出品エラー", selection.monsterId, error);
          selection.li.classList.remove("is-selected");
          const checkbox = selection.li.querySelector("input.es-bazaar-check");
          if (checkbox) checkbox.checked = false;
          controls.status.textContent = `${selection.name}: ${error.message}`;
        }
        if (index + 1 < selections.length) await sleep(REQUEST_DELAY_MS);
      }
      controls.status.textContent = errors ? `完了（成功:${selections.length - errors}、エラー:${errors}）` : "処理完了";
    } finally {
      controls.listButton.disabled = false;
      controls.selectAllButton.disabled = false;
      controls.multiplierInput.disabled = false;
      controls.renameInput.disabled = false;
      controls.closeButton.style.display = "";
    }
  };

  const init = () => {
    if (location.origin !== "https://eldersign.jp" || location.pathname !== "/book" || new URLSearchParams(location.search).get("cmd") !== "m") {
      alert("ブックのモンスター一覧画面で実行してください。");
      return;
    }
    const list = document.querySelector("nav.block ul");
    const heading = document.querySelector("header.page h1");
    if (!list || heading?.textContent.trim() !== "ブック") {
      alert("ブックのモンスター一覧を取得できませんでした。");
      return;
    }
    const controls = buildActionPanel();
    setupSelectableList(list, controls.status);
    let nextSelectAll = true;
    controls.selectAllButton.addEventListener("click", () => {
      setAllSelections(nextSelectAll, controls.status);
      controls.selectAllButton.textContent = nextSelectAll ? "全削除" : "全選択";
      nextSelectAll = !nextSelectAll;
    });
    controls.listButton.addEventListener("click", () => {
      runListing(controls).catch((error) => {
        console.error("一括バザー出品エラー", error);
        controls.status.textContent = `エラー: ${error.message}`;
        alert(`一括バザー出品でエラーが発生しました: ${error.message}`);
      });
    });
    const close = () => teardown(list);
    controls.closeButton.addEventListener("click", close);
    controls.closeButton.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        close();
      }
    });
  };

  try {
    init();
  } catch (error) {
    console.error("一括バザー出品エラー", error);
    alert(`一括バザー出品でエラーが発生しました: ${error.message}`);
  }
})();
