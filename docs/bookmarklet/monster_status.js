(function () {
  const RARITY = Object.freeze({
    BRONZE: "BRONZE",
    SILVER: "SILVER",
    GOLD: "GOLD",
    PLATINUM: "PLATINUM",
  });

  const RARITY_BY_MAX_LEVEL = Object.freeze({
    30: RARITY.BRONZE,
    50: RARITY.SILVER,
    70: RARITY.GOLD,
    90: RARITY.PLATINUM,
  });

  const GROWTH_COEFF_BY_RARITY = Object.freeze({
    [RARITY.BRONZE]: 1.0,
    [RARITY.SILVER]: 1.5,
    [RARITY.GOLD]: 2.0,
    [RARITY.PLATINUM]: 2.5,
  });

  const EXP_COEFF_BY_RARITY = Object.freeze({
    [RARITY.BRONZE]: 1,
    [RARITY.SILVER]: 2,
    [RARITY.GOLD]: 4,
    [RARITY.PLATINUM]: 8,
  });

  const FAMILY_NAME_BY_ID = Object.freeze({
    1: "妖精",
    2: "魔獣",
    3: "外来",
    4: "人造",
  });

  const TARGETS = ["HP", "攻撃", "魔力", "防御", "命中", "敏捷"];
  const GRADE_TOOL_URL = "https://yuki-kamikita.github.io/eldersign-tool/web/grade.html";
  // const GRADE_TOOL_URL = "http://localhost:8080/docs/web/grade.html";

  function getLevelInfo() {
    const h3 =
      document.querySelector("div.card_d header.card h3") ||
      document.querySelector("h3");
    if (!h3) {
      alert("レベル情報が見つかりません");
      return null;
    }

    const mLv = h3.textContent.match(/Lv\s*(\d+)\s*\/\s*(\d+)/i);
    if (!mLv) {
      alert("Lv/最大Lv形式が見つかりません");
      return null;
    }

    return { level: parseInt(mLv[1], 10), maxLevel: parseInt(mLv[2], 10) };
  }

  function getMonsterFamilyName() {
    const h3 =
      document.querySelector("div.card_d header.card h3") ||
      document.querySelector("h3");
    if (!h3) return null;

    const icon = h3.querySelector('img[src*="/img/menu/fm_"]');
    if (!icon || !icon.src) return null;

    const match = icon.src.match(/fm_(\d+)\.png/i);
    if (!match) return null;

    return FAMILY_NAME_BY_ID[parseInt(match[1], 10)] ?? null;
  }

  function getRarity(maxLevel) {
    return RARITY_BY_MAX_LEVEL[maxLevel] ?? RARITY.BRONZE;
  }

  function getGrowth(level, maxLevel, rarity) {
    const rc = GROWTH_COEFF_BY_RARITY[rarity] ?? 1.0;
    const t = Math.max(0, (level - 1) / (maxLevel - 1));
    const G = 1 + rc * t;
    return { G, sqrtG: Math.sqrt(G) };
  }

  function estimateLv1Base(currentTotal, bonus, isHP, level, G, sqrtG) {
    if (currentTotal <= 0) return 0;
    if (level <= 1) return Math.max(1, currentTotal - bonus);

    const factor = isHP ? sqrtG : G;
    if (!isFinite(factor) || factor <= 0) return Math.max(1, currentTotal - bonus);

    const approx = currentTotal / factor - bonus;

    let best = Math.max(1, Math.floor(approx));
    let bestDiff = Math.abs(Math.floor((best + bonus) * factor) - currentTotal);
    let maxExact = null;

    const start = Math.max(1, Math.floor(approx) - 50);
    const end = Math.floor(approx) + 50;

    for (let cand = start; cand <= end; cand++) {
      const sim = Math.floor((cand + bonus) * factor);
      const diff = Math.abs(sim - currentTotal);

      if (sim === currentTotal) {
        if (maxExact === null || cand > maxExact) maxExact = cand;
        continue;
      }
      if (diff < bestDiff) {
        bestDiff = diff;
        best = cand;
      }
    }
    return maxExact !== null ? maxExact : best;
  }

  function getStatusRows() {
    const cap = [...document.querySelectorAll("div.status table caption")].find(
      (c) => c.textContent.trim() === "ステータス"
    );
    if (!cap) {
      alert("モンスター画面で使用してください");
      return null;
    }
    return [...cap.closest("table").querySelectorAll("tbody tr")];
  }

  function collectStats(rows, level, G, sqrtG) {
    const lines = [];
    let sumSq = 0;
    const statInfo = {};

    rows.forEach((tr) => {
      const th = tr.querySelector("th");
      if (!th) return;

      const label = th.textContent.trim();
      if (!TARGETS.includes(label)) return;

      const isHP = label === "HP";
      const name = isHP ? "ＨＰ" : label;

      const tds = tr.querySelectorAll("td");
      if (tds.length < 2) return;

      const statText = tds[0].textContent.trim();
      const bonusText = tds[1].textContent.trim();

      const parts = statText.split("/");
      const usedPart = isHP && parts.length >= 2 ? parts[1] : parts[0];
      const currentTotal = parseInt(usedPart.replace(/[^\d\-]/g, ""), 10);
      if (isNaN(currentTotal)) return;

      const m = bonusText.match(/([+-]?\d+)/);
      const bonus = m ? parseInt(m[1], 10) : 0;

      const base = estimateLv1Base(currentTotal, bonus, isHP, level, G, sqrtG);
      if (base <= 0) return;

      const r = bonus / base;
      const pct = r * 100;

      const pctStr = String(pct.toFixed(1)).padStart(4, " ");
      const basePad = " ".repeat(Math.max(0, 5 - String(base).length));
      const bonusPad = " ".repeat(Math.max(0, 4 - String(bonus).length));

      lines.push(`${name}:${basePad}${base}+${bonusPad}${bonus} (${pctStr}%)`);

      sumSq += r * r;
      statInfo[label] = { name, base, bonus, r, current: currentTotal };
    });

    return { lines, sumSq, statInfo };
  }

  function calcEval(sumSq) {
    const raw = Math.sqrt(sumSq / 6) * 200 + 10;
    return Math.floor(raw * 10) / 10;
  }

  function getSourceRarityForLatent(rarity, monsterKind) {
    const isSubspecies = typeof monsterKind === "string" && monsterKind.startsWith("亜種");
    if (!isSubspecies) return rarity;
    if (rarity === RARITY.PLATINUM) return RARITY.GOLD;
    if (rarity === RARITY.GOLD) return RARITY.SILVER;
    if (rarity === RARITY.SILVER) return RARITY.BRONZE;
    return RARITY.BRONZE;
  }

  function getLatentSkillAlpha(rarity, monsterKind) {
    const sourceRarity = getSourceRarityForLatent(rarity, monsterKind);
    if (sourceRarity === RARITY.BRONZE) return 5;
    if (sourceRarity === RARITY.SILVER) return 2;
    return 0;
  }

  function calcLatentSkillRate(level, maxLevel, rarity, monsterKind) {
    if (!Number.isFinite(level) || !Number.isFinite(maxLevel) || maxLevel <= 0) {
      return null;
    }
    const alpha = getLatentSkillAlpha(rarity, monsterKind);
    return (level / maxLevel) * 100 + alpha;
  }

  function formatTrunc1(value) {
    if (!isFinite(value)) return "";
    const trunc = Math.trunc(value * 10) / 10;
    if (Math.abs(trunc - Math.round(trunc)) < 1e-9) return String(Math.round(trunc));
    return trunc.toFixed(1);
  }

  function formatNumber(value) {
    if (!isFinite(value)) return "";
    if (Math.abs(value - Math.round(value)) < 1e-9) return String(Math.round(value));
    return value.toFixed(1);
  }

  function getBazaarPriceAny() {
    const priceNode = document.querySelector("article.checktxt em");
    if (priceNode) {
      const m = priceNode.textContent.match(/([\d,]+)\s*Any/i);
      if (m) return parseInt(m[1].replace(/,/g, ""), 10);
    }

    const listingNode = document.querySelector("p.belt");
    if (listingNode) {
      const m = listingNode.textContent.match(/バザーに\s*([\d,]+)\s*Anyで出品中です/i);
      if (m) return parseInt(m[1].replace(/,/g, ""), 10);
    }

    const allText = document.body ? document.body.textContent : "";
    const textMatch = allText && allText.match(/出品金額は\s*([\d,]+)\s*Any/i);
    if (textMatch) return parseInt(textMatch[1].replace(/,/g, ""), 10);
    const listingTextMatch = allText && allText.match(/バザーに\s*([\d,]+)\s*Anyで出品中です/i);
    if (listingTextMatch) return parseInt(listingTextMatch[1].replace(/,/g, ""), 10);

    const scriptText = [...document.querySelectorAll("script")]
      .map((s) => s.textContent || "")
      .join("\n");
    const scriptMatch = scriptText.match(/confirm\.set\([^)]*?"([\d,]+)\s*Anyで購入します/i);
    if (scriptMatch) return parseInt(scriptMatch[1].replace(/,/g, ""), 10);

    return null;
  }

  function getGrade() {
    const gradeImg = document.querySelector('img[src*="/img/menu/grade_"]');
    if (!gradeImg) return null;
    const gm = gradeImg.src.match(/grade_(\d+)\.png/i);
    return gm ? parseInt(gm[1], 10) : null;
  }

  function getMonsterName() {
    const h2 =
      document.querySelector("div.card_d header.card h2") ||
      document.querySelector("h2");
    if (!h2) return null;
    return h2.textContent.replace(/\s+/g, " ").trim();
  }

  function getMonsterKind() {
    const card = document.querySelector("img#card");
    if (!card || !card.src) return null;
    const m = card.src.match(/_([0-9])g\d+/i);
    if (!m) return null;
    switch (m[1]) {
      case "0":
        return "原種";
      case "1":
        return "亜種1";
      case "2":
        return "亜種2";
      case "3":
        return "亜種3";
      default:
        return null;
    }
  }

  function getMonsterId() {
    const params = new URLSearchParams(window.location.search);
    const mid = params.get("mid");
    if (mid) return mid;

    const renameLink = document.querySelector('a[href*="/mcard_name?mid="]');
    if (!renameLink) return null;

    try {
      const url = new URL(renameLink.href, window.location.href);
      return url.searchParams.get("mid");
    } catch (_) {
      return null;
    }
  }

  async function renameMonster(monsterId, nextName) {
    if (!monsterId) throw new Error("モンスターIDが取得できません");
    const renameUrl = new URL("/mcard_name", window.location.origin);
    renameUrl.searchParams.set("mid", monsterId);
    renameUrl.searchParams.set("pg", "0");

    const renamePageResponse = await fetch(renameUrl.toString(), {
      credentials: "include",
    });
    if (!renamePageResponse.ok) {
      throw new Error("命名変更画面の取得に失敗しました");
    }

    const renamePageHtml = await renamePageResponse.text();
    const doc = new DOMParser().parseFromString(renamePageHtml, "text/html");
    const form = doc.querySelector('form[action*="mcard_detail"]');
    if (!form) {
      throw new Error("命名変更フォームが見つかりません");
    }

    const action = form.getAttribute("action");
    if (!action) {
      throw new Error("命名変更先URLが見つかりません");
    }

    const submitUrl = new URL(action, window.location.origin);
    const body = new URLSearchParams();
    body.set("name", nextName);

    const submitResponse = await fetch(submitUrl.toString(), {
      method: "POST",
      credentials: "include",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    });

    if (!submitResponse.ok) {
      throw new Error("命名変更の送信に失敗しました");
    }

    return nextName;
  }

  function buildGradeToolUrl(statInfo) {
    if (!statInfo) return null;
    const params = new URLSearchParams();

    const name = getMonsterName();
    const kind = getMonsterKind();
    if (name) params.set("name", name);
    if (kind) params.set("kind", kind);

    const keyMap = {
      HP: "hp",
      攻撃: "atk",
      魔力: "mag",
      防御: "def",
      命中: "hit",
      敏捷: "agi",
    };

    for (const label of TARGETS) {
      const info = statInfo[label];
      if (!info) continue;
      const key = keyMap[label];
      if (!key) continue;
      params.set(`${key}_base`, String(info.base));
      params.set(`${key}_bonus`, String(info.bonus));
    }

    const query = params.toString();
    return query ? `${GRADE_TOOL_URL}?${query}` : GRADE_TOOL_URL;
  }

  function buildMaxLevelLines(statInfo, maxLevel, rarity) {
    const rc = GROWTH_COEFF_BY_RARITY[rarity] ?? 1.0;
    const Gmax = 1 + rc;
    const sqrtGmax = Math.sqrt(Gmax);

    const pad5 = (n) => " ".repeat(Math.max(0, 5 - String(n).length));
    const pad4 = (n) => " ".repeat(Math.max(0, 4 - String(n).length));

    const lines = [];
    let maxAgi = null;
    let maxHit = null;
    lines.push("-----------------------");
    lines.push(`Lv1->${maxLevel}(最大)時のステータス`);
    for (const label of TARGETS) {
      const info = statInfo[label];
      if (!info) continue;

      const factor = label === "HP" ? sqrtGmax : Gmax;
      const lv1status = info.base + info.bonus
      const maxTotal = Math.floor(lv1status * factor);
      if (label === "敏捷") maxAgi = maxTotal;
      if (label === "命中") maxHit = maxTotal;

      lines.push(
        `${info.name}:${pad5(lv1status)}${lv1status} ->${pad5(maxTotal)}${maxTotal}`
      );
    }
    if (maxHit != null) {
      const maxHitRate = Math.sqrt(maxHit);
      lines.push(`命中率: ${formatNumber(maxHitRate)}`);
    }
    if (maxAgi != null) {
      const maxEvasionRate = Math.sqrt(maxAgi);
      lines.push(`回避率: ${formatNumber(maxEvasionRate)}`);
      const maxActionValue = Math.sqrt(maxAgi * 2000);
      lines.push(`行動値: ${Math.round(maxActionValue)}`);
    }

    return lines;
  }

  function buildNextGradeLines(statInfo, evalValue) {
    const currentTens = Math.floor(evalValue / 10);
    const target = (currentTens + 1) * 10;

    function rawEval(sumR2) {
      return Math.sqrt(sumR2 / 6) * 200 + 10;
    }

    function sumR2Override(label, newB) {
      let s = 0;
      for (const k of TARGETS) {
        const info = statInfo[k];
        if (!info) continue;
        const b = k === label ? newB : info.bonus;
        const r = b / info.base;
        s += r * r;
      }
      return s;
    }

    function findMin(label) {
      const info = statInfo[label];
      if (!info) return null;

      let S = 0;
      for (const k of TARGETS) {
        if (k === label) continue;
        const o = statInfo[k];
        if (o) S += o.r * o.r;
      }

      const k = (target - 10) / 200;
      const need = Math.max(0, 6 * k * k - S);
      let start = Math.max(info.bonus, Math.ceil(Math.sqrt(need) * info.base) - 3);

      for (let b = start; b <= info.bonus + 20000; b++) {
        if (rawEval(sumR2Override(label, b)) >= target) return b;
      }
      return null;
    }

    const pad4 = (n) => String(n).padStart(4, " ");
    const pad3 = (n) => String(n).padStart(3, " ");
    const padPct = (x) => String(x.toFixed(1)).padStart(4, " ");

    const lines = [];
    lines.push("-----------------------");
    lines.push("次のグレードになるには");
    for (const label of TARGETS) {
      const info = statInfo[label];
      if (!info) continue;

      const b = findMin(label);
      if (b == null) continue;

      const pct = (b / info.base) * 100;
      const delta = b - info.bonus;

      lines.push(`${info.name}: +${pad4(b)}(${padPct(pct)}%) あと${pad3(delta)}`);
    }

    return lines;
  }

  function buildCombatMemoLines() {
    return [
      "-----------------------",
      "命中と回避の差4.5≒命中率10%",
    ];
  }

  function renderPanel(lines, moreLines, gradeUrl, panelTitle, options = {}) {
    const { renameActions = {} } = options;
    const overlayPadTop = 10;
    const overlayPadBottom = 10;
    const overlay = document.createElement("div");
    overlay.style.cssText =
      "position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.45);" +
      `display:flex;justify-content:flex-end;align-items:flex-start;padding:${overlayPadTop}px 10px ${overlayPadBottom}px 10px;`;

    const box = document.createElement("div");
    box.style.cssText =
      "background:rgba(0,0,0,.6);color:#fff;border-radius:8px;font-family:monospace;" +
      "font-size:14px;width:fit-content;max-width:min(92vw,420px);display:flex;flex-direction:column;" +
      "box-shadow:0 8px 24px rgba(0,0,0,.35);position:relative;";
    overlay.appendChild(box);

    const header = document.createElement("div");
    header.style.cssText =
      "display:flex;justify-content:space-between;align-items:center;" +
      "padding:8px 10px;border-bottom:1px solid rgba(255,255,255,.2);";
    box.appendChild(header);

    const title = document.createElement("div");
    title.textContent = panelTitle || "モンスター情報";
    title.style.cssText = "";
    header.appendChild(title);

    const closeBtn = document.createElement("span");
    closeBtn.setAttribute("role", "button");
    closeBtn.setAttribute("tabindex", "0");
    closeBtn.title = "閉じる";
    closeBtn.style.cssText =
      "position:absolute;top:0px;right:0px;cursor:pointer;" +
      "color:#fff;font-size:16px;line-height:16px;padding:6px 10px;";
    closeBtn.innerHTML =
      '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" ' +
      'fill="currentColor"><path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>';
    box.appendChild(closeBtn);

    const content = document.createElement("div");
    content.style.cssText =
      "padding:10px 12px;overflow-y:auto;overflow-x:hidden;" +
      "-webkit-overflow-scrolling:touch;touch-action:pan-y;overscroll-behavior:contain;";
    box.appendChild(content);

    function applyModalSize() {
      const viewportH =
        (window.visualViewport && Number.isFinite(window.visualViewport.height)
          ? window.visualViewport.height
          : window.innerHeight) || 0;
      const maxH = Math.max(180, Math.floor(viewportH - overlayPadTop - overlayPadBottom));
      box.style.maxHeight = `${maxH}px`;
    }
    applyModalSize();
    window.addEventListener("resize", applyModalSize);

    function appendLines(target, list) {
      list.forEach((line, index) => {
        if (index > 0) target.appendChild(document.createTextNode("\n"));
        if (index === 0 && /^\S.*\((原種|亜種[1-3])\)$/.test(line)) {
          const centered = document.createElement("span");
          centered.textContent = line;
          centered.style.cssText = "display:block;text-align:center;";
          target.appendChild(centered);
          return;
        }
        if (gradeUrl && line === "次のグレードになるには") {
          const link = document.createElement("a");
          link.href = gradeUrl;
          link.target = "_blank";
          link.rel = "noopener";
          link.textContent = line;
          link.style.cssText = "text-decoration:underline;color:inherit;cursor:pointer;";
          link.onclick = (e) => {
            e.stopPropagation();
          };
          target.appendChild(link);
          return;
        }
        const renameAction =
          (/^納品pt:\s*\d+/.test(line) && renameActions.delivery) ||
          (/^[^\s]+経験値:\s*\d+/.test(line) && renameActions.experience) ||
          null;
        if (renameAction) {
          const wrap = document.createElement("span");
          wrap.textContent = line;

          const renameBtn = document.createElement("a");
          renameBtn.href = "#";
          renameBtn.textContent = "命名";
          renameBtn.style.cssText =
            "margin-left:8px;color:inherit;" +
            "font-size:12px;cursor:pointer;text-decoration:underline;";
          renameBtn.onclick = async (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (renameBtn.disabled) return;
            renameBtn.disabled = true;
            const originalText = renameBtn.textContent;
            renameBtn.textContent = "変更中";
            try {
              await renameAction();
              renameBtn.textContent = "変更済み";
            } catch (error) {
              renameBtn.disabled = false;
              renameBtn.textContent = originalText;
              alert("命名変更に失敗しました: " + error.message);
            }
          };

          wrap.appendChild(renameBtn);
          target.appendChild(wrap);
          return;
        }
        target.appendChild(document.createTextNode(line));
      });
    }

    const pre = document.createElement("pre");
    pre.style.cssText = "margin:0;white-space:pre-wrap;word-break:break-word;";
    appendLines(pre, lines);
    content.appendChild(pre);

    const toggleWrap = document.createElement("div");
    toggleWrap.style.cssText = "margin-top:8px;text-align:right;padding-bottom:4px;";
    content.appendChild(toggleWrap);

    const toggle = document.createElement("a");
    toggle.href = "#";
    toggle.textContent = "もっと見る";
    toggle.style.cssText = "font-size:12px;cursor:pointer;text-decoration:underline;color:inherit;";
    toggleWrap.appendChild(toggle);

    let opened = false;
    toggle.onclick = (e) => {
      e.preventDefault();
      if (opened) return;
      const nextPre = document.createElement("pre");
      nextPre.style.cssText = "margin:8px 0 0;white-space:pre-wrap;word-break:break-word;";
      appendLines(nextPre, moreLines);
      content.appendChild(nextPre);
      toggleWrap.remove();
      opened = true;
    };

    function closePanel() {
      window.removeEventListener("resize", applyModalSize);
      overlay.remove();
    }

    closeBtn.onclick = (e) => {
      e.stopPropagation();
      closePanel();
    };
    closeBtn.onkeydown = (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.stopPropagation();
        closePanel();
      }
    };
    overlay.onclick = (e) => {
      if (e.target !== overlay) return;
      closePanel();
    };
    document.body.appendChild(overlay);
  }

  try {
    const levelInfo = getLevelInfo();
    if (!levelInfo) return;

    const { level, maxLevel } = levelInfo;
    const rarity = getRarity(maxLevel);
    const { G, sqrtG } = getGrowth(level, maxLevel, rarity);
    const expRarityFactor = EXP_COEFF_BY_RARITY[rarity] ?? 1;
    const monsterId = getMonsterId();
    const familyName = getMonsterFamilyName();

    const rows = getStatusRows();
    if (!rows) return;

    const { lines, sumSq, statInfo } = collectStats(rows, level, G, sqrtG);
    const evalValue = calcEval(sumSq);
    const grade = getGrade();
    const monsterName = getMonsterName();
    const monsterKind = getMonsterKind();
    const rarityCoeff = GROWTH_COEFF_BY_RARITY[rarity] ?? 1.0;
    const deliveryPoint = evalValue * rarityCoeff;
    const deliveryPointDisplay = Math.floor(deliveryPoint);
    const bazaarPriceAny = getBazaarPriceAny();

    const levelSuffix = Number.isFinite(level) ? ` Lv${level}` : "";
    const panelTitle = monsterName
      ? `${monsterName}${monsterKind ? `(${monsterKind})` : ""}${levelSuffix}`
      : `モンスター情報${levelSuffix}`;

    const agiInfo = statInfo["敏捷"];
    const currentStatusTailLines = [];
    lines.push("評価値: " + evalValue.toFixed(1));
    if (bazaarPriceAny != null && deliveryPointDisplay > 0) {
      const unitPrice = bazaarPriceAny / deliveryPoint;
      const unitPriceRounded = Math.round(unitPrice);
      lines.push(`納品pt: ${deliveryPointDisplay} (${unitPriceRounded}any/pt)`);
    } else {
      lines.push(`納品pt: ${deliveryPointDisplay}`);
    }
    lines.push("-----------------------");
    lines.push("現在ステータス準拠");
    const hitInfo = statInfo["命中"];
    if (hitInfo) {
      const hitRate = Math.sqrt(hitInfo.current);
      lines.push(`命中率: ${formatNumber(hitRate)}`);
    }
    if (agiInfo) {
      const evasionRate = Math.sqrt(agiInfo.current);
      lines.push(`回避率: ${formatNumber(evasionRate)}`);
      const agility = agiInfo.current;
      const actionValue = Math.sqrt(agility * 2000);
      lines.push(`行動値: ${Math.round(actionValue)}`);
      if (grade != null) {
        const baseExp = expRarityFactor * grade * ((level + 4) / 5) * 16;
        const expS = Math.floor(baseExp * 1.125);
        currentStatusTailLines.push(`${familyName || "同族"}経験値: ${expS}`);
      } else {
        currentStatusTailLines.push("経験値: 取得失敗");
      }
      const latentRateCurrent = calcLatentSkillRate(level, maxLevel, rarity, monsterKind);
      if (latentRateCurrent != null) {
        currentStatusTailLines.push(`潜在発現率: ${formatTrunc1(latentRateCurrent)}%`);
      }
    }
    lines.push(...currentStatusTailLines);

    const maxLevelLines = buildMaxLevelLines(statInfo, maxLevel, rarity);
    const nextGradeLines = buildNextGradeLines(statInfo, evalValue);
    const combatMemoLines = buildCombatMemoLines();
    const moreLines = [...maxLevelLines, ...nextGradeLines, ...combatMemoLines];
    const gradeUrl = buildGradeToolUrl(statInfo);
    renderPanel(lines, moreLines, gradeUrl, panelTitle, {
      renameActions: {
        delivery: async () => {
          const nextName = await renameMonster(monsterId, `納品基礎pt${deliveryPointDisplay}`);
          const heading =
            document.querySelector("div.card_d header.card h1") ||
            document.querySelector("h1");
          if (heading) {
            const textNode = [...heading.childNodes].find((node) => node.nodeType === Node.TEXT_NODE);
            if (textNode) {
              textNode.textContent = nextName + " ";
            } else {
              heading.prepend(document.createTextNode(nextName + " "));
            }
          }
        },
        experience: async () => {
          if (grade == null) {
            throw new Error("経験値が取得できません");
          }
          const baseExp = expRarityFactor * grade * ((level + 4) / 5) * 16;
          const expS = Math.floor(baseExp * 1.125);
          const nextName = await renameMonster(monsterId, `${familyName || "同族"}経験値${expS}`);
          const heading =
            document.querySelector("div.card_d header.card h1") ||
            document.querySelector("h1");
          if (heading) {
            const textNode = [...heading.childNodes].find((node) => node.nodeType === Node.TEXT_NODE);
            if (textNode) {
              textNode.textContent = nextName + " ";
            } else {
              heading.prepend(document.createTextNode(nextName + " "));
            }
          }
        },
      },
    });
  } catch (e) {
    alert("エラー: " + e.message);
  }
})();
