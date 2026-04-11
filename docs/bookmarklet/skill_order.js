(() => {
  const PANEL_ID = "__es_skill_order_panel";
  const DEFAULT_LEFT = "左";
  const DEFAULT_RIGHT = "右";
  const GLOBAL_API = "ElderSignSkillOrder";

  function normalizeName(name) {
    return name.replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
  }

  function getLinesFromElement(el) {
    const lines = [];
    let current = "";

    const flush = () => {
      const text = current.trim();
      if (text) lines.push(text);
      current = "";
    };

    el.childNodes.forEach((node) => {
      if (node.nodeType === Node.TEXT_NODE) {
        current += node.textContent || "";
        return;
      }
      if (node.nodeName === "BR") {
        flush();
        return;
      }
      current += node.textContent || "";
    });

    flush();
    return lines;
  }

  function parsePartyCell(td) {
    const monsters = [];
    const ps = td.querySelectorAll("p");
    ps.forEach((p) => {
      const lines = getLinesFromElement(p);
      if (!lines.length) return;
      const nameLine = lines[0];
      const name = parseNameLine(nameLine);
      const level = parseLevel(nameLine);
      const statusLine = lines.length >= 2 ? lines[1].replace(/\s+/g, " ").trim() : "";
      const status = parseInitialStatus(statusLine);
      monsters.push({ name, level, status });
    });
    return monsters;
  }

  function parseNameLine(line) {
    const nameMatch = line.match(/^(.*)\s+Lv\d+/);
    return normalizeName(nameMatch ? nameMatch[1] : line);
  }

  function parseLevel(line) {
    const m = line.match(/Lv\s*(\d+)/i);
    return m ? parseInt(m[1], 10) : null;
  }

  function parseInitialStatus(line) {
    if (!line) return "";
    const m = line.match(/(\d+)\s*\/\s*(\d+)\s*(.*)$/);
    if (!m) return "";
    const maxHp = m[2];
    const state = normalizeName(m[3] || "");
    return `HP${maxHp}${state ? " " + state : ""}`;
  }

  function buildMonsterMap(list) {
    const order = [];
    const map = new Map();
    list.forEach((m) => {
      if (map.has(m.name)) return;
      order.push(m.name);
      map.set(m.name, {
        name: m.name,
        level: m.level ?? null,
        status: m.status,
        turns: {},
        turnDetails: {},
      });
    });
    return { order, map };
  }

  function ensureMonster(mapInfo, name) {
    if (mapInfo.map.has(name)) return mapInfo.map.get(name);
    mapInfo.order.push(name);
    const m = { name, level: null, status: "", turns: {}, turnDetails: {} };
    mapInfo.map.set(name, m);
    return m;
  }

  function addSkill(mapInfo, name, turn, skill, detail = {}) {
    if (!skill) return;
    const m = ensureMonster(mapInfo, name);
    if (!m.turns[turn]) m.turns[turn] = [];
    m.turns[turn].push(skill);
    if (!m.turnDetails[turn]) m.turnDetails[turn] = [];
    m.turnDetails[turn].push({
      skill,
      variant: detail.variant ?? null,
      imageUrl: detail.imageUrl || "",
    });
  }

  function parseTurnNumber(section) {
    const h4 = section.querySelector("header h4");
    if (!h4) return null;
    const m = h4.textContent.match(/Turn\s*(\d+)/i);
    return m ? parseInt(m[1], 10) : null;
  }

  function extractSkillFromEm(text) {
    const t = text.trim();
    const attackMatch = t.match(/^(.+?)の攻撃！(.*)$/);
    if (attackMatch) {
      const actor = normalizeName(attackMatch[1]);
      let skill = attackMatch[2].trim();
      if (skill.endsWith("！")) skill = skill.slice(0, -1).trim();
      if (!skill) return { actor, skill: null };
      return { actor, skill };
    }
    const generalMatch = t.match(/^(.+?)の.+?！(.*)$/);
    if (generalMatch) {
      const actor = normalizeName(generalMatch[1]);
      let skill = generalMatch[2].trim();
      if (skill.endsWith("！")) skill = skill.slice(0, -1).trim();
      if (!skill) return { actor, skill: null };
      return { actor, skill };
    }
    return null;
  }

  function extractSupportSkill(text) {
    const t = text.trim();
    const m = t.match(/！([^！]+)！/);
    if (m) return m[1].trim();
    const parts = t.split("！");
    if (parts.length >= 2) return parts[1].trim();
    return null;
  }

  function isSupportEffectLine(line) {
    if (!line) return false;
    if (!line.includes("は")) return false;
    if (/(効果が無かった|抵抗した|回避した|なんともない|倒れた|解けた)/.test(line)) return false;
    if (/ダメージ！/.test(line)) return false;
    return /！$/.test(line);
  }

  function renderPanel(leftTitle, rightTitle, leftText, rightText) {
    document.getElementById(PANEL_ID)?.remove();

    const box = document.createElement("div");
    box.id = PANEL_ID;
    const boxTop = 10;
    box.style.cssText =
      `position:fixed;top:${boxTop}px;right:10px;z-index:99999;` +
      "background:rgba(0,0,0,.8);color:#fff;padding:24px 12px 12px;" +
      "border-radius:8px;font-family:monospace;font-size:13px;" +
      "max-width:calc(100% - 20px);max-height:calc(100% - 20px);" +
      "overflow:auto;display:inline-flex;" +
      "flex-direction:column;align-items:flex-end;";

    const closePanel = () => {
      window.removeEventListener("resize", onResize);
      if (window.visualViewport) {
        window.visualViewport.removeEventListener("resize", onResize);
      }
      box.remove();
    };

    const close = document.createElement("span");
    close.setAttribute("role", "button");
    close.setAttribute("tabindex", "0");
    close.title = "閉じる";
    close.style.cssText =
      "position:absolute;top:0px;right:0px;cursor:pointer;" +
      "color:#fff;font-size:16px;line-height:16px;padding:6px 10px;";
    close.innerHTML =
      '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" ' +
      'fill="currentColor"><path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>';
    close.onclick = (e) => {
      e.stopPropagation();
      closePanel();
    };
    close.onkeydown = (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.stopPropagation();
        closePanel();
      }
    };

    const row = document.createElement("div");
    row.style.cssText =
      "display:flex;gap:12px;flex-wrap:nowrap;justify-content:flex-end;";

    const panels = [];

    const setExpanded = (item, expanded) => {
      item.panel.dataset.expanded = expanded ? "1" : "0";
      item.toggle.title = expanded ? "閉じる" : "開く";
      item.toggle.style.transform = expanded ? "rotate(90deg)" : "rotate(0deg)";
      item.body.style.opacity = expanded ? "1" : "0";
      item.body.style.overflow = expanded ? "auto" : "hidden";
      const maxHeight = item.body.dataset.maxHeight || "0";
      item.body.style.maxHeight = expanded ? maxHeight : "0px";
    };

    const makePanel = (title, text, isLeft) => {
      const panel = document.createElement("div");
      panel.style.cssText =
        "min-width:280px;max-width:46vw;border:1px solid rgba(255,255,255,.2);" +
        "border-radius:6px;padding:8px;display:flex;flex-direction:column;gap:6px;min-height:0;";

      const head = document.createElement("div");
      head.style.cssText =
        "display:flex;align-items:center;justify-content:space-between;gap:8px;cursor:pointer;";

      const leftWrap = document.createElement("div");
      leftWrap.style.cssText = "display:flex;align-items:center;gap:6px;";

      const toggle = document.createElement("span");
      toggle.setAttribute("role", "button");
      toggle.setAttribute("tabindex", "0");
      toggle.title = "開く";
      toggle.style.cssText =
        "display:none;cursor:pointer;user-select:none;line-height:1;" +
        "width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;" +
        "transition:transform 0.18s ease;";
      toggle.innerHTML =
        '<svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" ' +
        'fill="currentColor"><path d="M8.59 16.59 13.17 12 8.59 7.41 10 6l6 6-6 6z"/></svg>';

      const label = document.createElement("div");
      label.textContent = title;
      label.style.cssText = "font-weight:bold;";

      const btn = document.createElement("span");
      btn.title = "コピー";
      btn.setAttribute("role", "button");
      btn.setAttribute("tabindex", "0");
      btn.style.cssText =
        "cursor:pointer;background:transparent;color:#fff;border:1px solid transparent;" +
        "border-radius:4px;padding:2px;line-height:0;display:flex;align-items:center;";
      btn.innerHTML =
        '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"' +
        ' fill="currentColor"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1z"/>' +
        '<path d="M20 5H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h12v14z"/></svg>';

      const onCopy = async (e) => {
        e.stopPropagation();
        await copyText(wrapCopyText(title, text));
        btn.style.color = "#9fd";
        setTimeout(() => {
          btn.style.color = "#fff";
        }, 1200);
      };
      btn.onclick = onCopy;
      btn.onkeydown = (e) => {
        if (e.key === "Enter" || e.key === " ") onCopy(e);
      };

      leftWrap.appendChild(toggle);
      leftWrap.appendChild(label);
      head.appendChild(leftWrap);
      head.appendChild(btn);

      const body = document.createElement("div");
      body.style.cssText =
        "overflow:auto;max-height:70vh;min-height:0;-webkit-overflow-scrolling:touch;" +
        "transition:max-height 0.25s ease, opacity 0.2s ease;";

      const pre = document.createElement("pre");
      pre.style.cssText = "margin:0;white-space:pre;color:#fff;";
      pre.textContent = text;

      panel.appendChild(head);
      body.appendChild(pre);
      panel.appendChild(body);
      row.appendChild(panel);
      const item = { panel, head, body, toggle };
      panels.push(item);

      const togglePanel = (e) => {
        e.stopPropagation();
        const next = panel.dataset.expanded !== "1";
        setExpanded(item, next);
      };
      toggle.onclick = togglePanel;
      toggle.onkeydown = (e) => {
        if (e.key === "Enter" || e.key === " ") togglePanel(e);
      };
      head.onclick = (e) => {
        if (btn.contains(e.target)) return;
        togglePanel(e);
      };
      head.onkeydown = (e) => {
        if (btn.contains(e.target)) return;
        if (e.key === "Enter" || e.key === " ") togglePanel(e);
      };
      head.setAttribute("role", "button");
      head.setAttribute("tabindex", "0");

    };

    makePanel(leftTitle, leftText, true);
    makePanel(rightTitle, rightText, false);

    const applyLayout = () => {
      row.style.flexDirection = "row";
      row.style.alignItems = "stretch";
      panels.forEach((item) => {
        const panel = item.panel;
        panel.style.minWidth = "280px";
        panel.style.maxWidth = "46vw";
        panel.style.width = "";
      });

      const available = Math.max(0, window.innerWidth - 20);
      const rowWidth = row.scrollWidth;
      const boxRect = box.getBoundingClientRect();
      if (rowWidth > available || boxRect.left < 10) {
        row.style.flexDirection = "column";
        row.style.alignItems = "flex-end";
        panels.forEach((item) => {
          const panel = item.panel;
          panel.style.minWidth = "0";
          panel.style.maxWidth = "calc(100vw - 20px - 24px)";
        });
        const widths = panels.map((item) => item.panel.getBoundingClientRect().width);
        const maxWidth = Math.min(available - 24, Math.max(...widths, 0));
        panels.forEach((item) => {
          item.panel.style.width = `${Math.ceil(maxWidth)}px`;
        });
      }

      const viewport = window.visualViewport;
      const viewportHeight = viewport ? viewport.height : window.innerHeight;
      const narrow = window.innerWidth < 520;
      const fallbackBottom = !viewport && narrow ? 64 : 0;
      const topOffset = Number.parseInt(box.style.top, 10) || boxTop;
      const maxHeight = Math.max(0, viewportHeight - topOffset - 32 - fallbackBottom);
      box.style.maxHeight = `${Math.floor(maxHeight)}px`;
      const needsAccordion = row.scrollHeight + 24 > maxHeight;
      const useAccordion = narrow || needsAccordion;
      panels.forEach((item) => {
        if (!useAccordion) {
          item.toggle.style.display = "none";
          item.body.dataset.maxHeight = "70vh";
          setExpanded(item, true);
          item.panel.dataset.accordion = "0";
          return;
        }
        item.toggle.style.display = "inline-block";
        if (item.panel.dataset.accordion !== "1") setExpanded(item, true);
        item.panel.dataset.accordion = "1";
        const headHeight = item.head.getBoundingClientRect().height;
        const maxBodyHeight = Math.max(80, maxHeight - headHeight - 60);
        item.body.dataset.maxHeight = `${Math.floor(maxBodyHeight)}px`;
        if (item.panel.dataset.expanded === "1") {
          item.body.style.maxHeight = item.body.dataset.maxHeight;
        }
      });
    };

    const onResize = () => applyLayout();

    box.appendChild(close);
    box.appendChild(row);
    document.body.appendChild(box);
    window.addEventListener("resize", onResize);
    if (window.visualViewport) {
      window.visualViewport.addEventListener("resize", onResize);
    }
    requestAnimationFrame(applyLayout);
  }

  async function copyText(text) {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return;
    }
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.top = "-1000px";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
  }

  function wrapCopyText(title, text) {
    return "```\n" + title + "\n" + text + "\n```";
  }

  function buildOutputText(mapInfo, maxTurn, presentByTurn) {
    const lines = [];
    mapInfo.order.forEach((name) => {
      const m = mapInfo.map.get(name);
      if (!m) return;
      const status = m.status || "-";
      const levelText = m.level != null ? `Lv${m.level}` : "";
      lines.push([m.name, levelText].filter(Boolean).join(" "));
      lines.push(status);

      const p0Skills = m.turns[0] || [];
      lines.push(` P0: ${p0Skills.length ? p0Skills.join(" / ") : "-"}`);

      const seen = new Set(p0Skills);
      const actions = [];
      for (let t = 1; t <= maxTurn; t++) {
        if (presentByTurn && !presentByTurn[t]?.has(name)) continue;
        const skills = m.turns[t] || [];
        skills.forEach((skill) => {
          if (!skill || seen.has(skill)) return;
          seen.add(skill);
          actions.push(skill);
        });
      }
      actions.forEach((skill, idx) => {
        lines.push(` A${idx + 1}: ${skill}`);
      });
      lines.push("");
    });
    if (lines.length && lines[lines.length - 1] === "") lines.pop();
    return lines.join("\n");
  }

  function getPartyTitles(doc = document) {
    const h1 = doc.querySelector("section.btl header.btl h1");
    if (!h1) return { left: DEFAULT_LEFT, right: DEFAULT_RIGHT };
    const parts = h1.textContent.split("vs");
    if (parts.length < 2) return { left: DEFAULT_LEFT, right: DEFAULT_RIGHT };
    return {
      left: normalizeName(parts[0]),
      right: normalizeName(parts[1]),
    };
  }

  function collectPresence(td) {
    const names = new Set();
    const ps = td.querySelectorAll("p");
    ps.forEach((p) => {
      const lines = getLinesFromElement(p);
      if (!lines.length) return;
      const name = parseNameLine(lines[0]);
      if (name) names.add(name);
    });
    return names;
  }

  function extractVariantNumber(src) {
    if (!src) return null;
    const m = String(src).match(/_(\d+)g\d+\.png/i);
    if (!m) return null;
    const value = parseInt(m[1], 10);
    return Number.isNaN(value) ? null : value;
  }

  function getImageFromRound(roundEl) {
    let node = roundEl.previousElementSibling;
    while (node) {
      if (node.tagName === "H5") {
        const img = node.querySelector("img");
        return img?.src || "";
      }
      if (node.tagName === "HEADER") break;
      node = node.previousElementSibling;
    }
    return "";
  }

  function parseBattleDocument(doc = document) {
    const turnSections = Array.from(doc.querySelectorAll("section.turn"));
    if (!turnSections.length) {
      throw new Error("バトル結果ページで使用してください");
    }

    const firstTurn = turnSections[0];
    const partyTable = firstTurn.querySelector("table.party");
    if (!partyTable) {
      throw new Error("初期陣営情報が見つかりません");
    }

    const partyCells = partyTable.querySelectorAll("tbody tr td");
    if (partyCells.length < 2) {
      throw new Error("陣営情報の取得に失敗しました");
    }

    const leftMonsters = parsePartyCell(partyCells[0]);
    const rightMonsters = parsePartyCell(partyCells[1]);

    const leftInfo = buildMonsterMap(leftMonsters);
    const rightInfo = buildMonsterMap(rightMonsters);

    const allTurns = [];
    turnSections.forEach((section) => {
      const t = parseTurnNumber(section);
      if (t != null) allTurns.push(t);
    });
    const maxTurn = allTurns.length ? Math.max(...allTurns) : 0;

    const presentLeftByTurn = [];
    const presentRightByTurn = [];
    presentLeftByTurn[0] = new Set(leftMonsters.map((m) => m.name));
    presentRightByTurn[0] = new Set(rightMonsters.map((m) => m.name));

    const preTurnRounds = Array.from(doc.querySelectorAll("p.round")).filter(
      (p) => !p.closest("section.turn")
    );
    preTurnRounds.forEach((p) => {
      const em = p.querySelector("em");
      if (!em) return;
      const skill = extractSupportSkill(em.textContent);
      if (!skill) return;

      const lines = getLinesFromElement(p);
      const added = new Set();
      lines.forEach((line) => {
        if (!isSupportEffectLine(line)) return;
        const m = line.match(/^(.+?)は(.+?)！$/);
        if (!m) return;
        const target = normalizeName(m[1]);
        if (added.has(target)) return;
        if (leftInfo.map.has(target)) {
          addSkill(leftInfo, target, 0, skill);
          added.add(target);
        } else if (rightInfo.map.has(target)) {
          addSkill(rightInfo, target, 0, skill);
          added.add(target);
        }
      });
    });

    turnSections.forEach((section) => {
      const turn = parseTurnNumber(section);
      if (turn == null) return;

      const table = section.querySelector("table.party");
      if (table) {
        const cells = table.querySelectorAll("tbody tr td");
        if (cells.length >= 2) {
          presentLeftByTurn[turn] = collectPresence(cells[0]);
          presentRightByTurn[turn] = collectPresence(cells[1]);
        }
      }

      const rounds = section.querySelectorAll("p.round");
      rounds.forEach((p) => {
        const imageUrl = getImageFromRound(p);
        const variant = extractVariantNumber(imageUrl);
        const ems = p.querySelectorAll("em");
        let parsed = null;
        for (const em of ems) {
          parsed = extractSkillFromEm(em.textContent);
          if (parsed) break;
        }
        if (!parsed || !parsed.actor) return;
        if (!parsed.skill) return;

        if (leftInfo.map.has(parsed.actor)) {
          addSkill(leftInfo, parsed.actor, turn, parsed.skill, { variant, imageUrl });
        } else if (rightInfo.map.has(parsed.actor)) {
          addSkill(rightInfo, parsed.actor, turn, parsed.skill, { variant, imageUrl });
        }
      });
    });

    return {
      leftInfo,
      rightInfo,
      maxTurn,
      presentLeftByTurn,
      presentRightByTurn,
      titles: getPartyTitles(doc),
    };
  }

  function parseBattleHtml(html) {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return parseBattleDocument(doc);
  }

  function run(doc = document) {
    const result = parseBattleDocument(doc);
    const leftText = buildOutputText(result.leftInfo, result.maxTurn, result.presentLeftByTurn);
    const rightText = buildOutputText(result.rightInfo, result.maxTurn, result.presentRightByTurn);
    renderPanel(result.titles.left, result.titles.right, leftText, rightText);
    return result;
  }

  window[GLOBAL_API] = {
    parseBattleDocument,
    parseBattleHtml,
    buildOutputText,
    run,
    normalizeName,
  };

  if (window.__ES_SKILL_ORDER_NO_AUTO_RUN) return;

  try {
    run(document);
  } catch (e) {
    alert("エラー: " + e.message);
  }
})();
