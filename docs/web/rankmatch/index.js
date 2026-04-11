let PHASES = [];

      const GROUP_ORDER = ["S", "A1", "A2"];

      const phaseTabs = document.getElementById("phaseTabs");
      const groupGrid = document.getElementById("groupGrid");
      const searchInput = document.getElementById("searchInput");
      const searchStatus = document.getElementById("searchStatus");
      const openAllPlayersButton = document.getElementById("openAllPlayers");
      const closeAllPlayersButton = document.getElementById("closeAllPlayers");
      const phaseJumpForm = document.getElementById("phaseJumpForm");
      const phaseJumpInput = document.getElementById("phaseJumpInput");
      let currentPhase = null;
      let currentQuery = "";
      let phaseMeta = [];
      let phaseByNumber = new Map();

      const renderMessage = (text) => {
        groupGrid.innerHTML = "";
        const empty = document.createElement("div");
        empty.className = "empty";
        empty.textContent = text;
        groupGrid.appendChild(empty);
      };

      const updateSearchStatus = (count) => {
        if (!searchStatus) return;
        if (!currentQuery) {
          searchStatus.textContent = "";
          return;
        }
        searchStatus.textContent = `ヒット: ${count}件`;
      };

      const setAllPlayerAccordions = (isOpen) => {
        document.querySelectorAll("details.player-accordion").forEach((detail) => {
          detail.open = isOpen;
        });
      };

      const loadPhaseMap = async () => {
        try {
          const response = await fetch("./phase_map.json", { cache: "no-store" });
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const data = await response.json();
          const phases = Array.isArray(data) ? data : data.phases;
          return Array.isArray(phases) ? phases : [];
        } catch (err) {
          console.warn("phase_map.jsonの読み込みに失敗しました。", err);
          return [];
        }
      };

      const extractPhaseNumber = (phase, fallback) => {
        const sources = [phase?.label, phase?.id, phase?.name];
        for (const source of sources) {
          if (!source) continue;
          const match = String(source).match(/(\d+)/);
          if (match) return parseInt(match[1], 10);
        }
        return fallback;
      };

      const buildPhaseMeta = (phases) => {
        return phases.map((phase, index) => ({
          phase,
          index,
          number: extractPhaseNumber(phase, index + 1),
        }));
      };

      const setActivePhase = (meta) => {
        if (!meta?.phase) return;
        currentPhase = meta.phase;
        renderPhase(currentPhase);
        document.querySelectorAll(".phase-button").forEach((tab) => {
          tab.classList.toggle("is-active", Number(tab.dataset.index) === meta.index);
        });
      };

      const LETTERS = "ABCDEFGHIJKL".split("");

      const buildSchedule = (letters) => {
        if (letters.length < 2) return {};
        const anchor = letters[0];
        const rest = letters.slice(1).reverse();
        const circle = [anchor, ...rest];
        const rounds = letters.length - 1;
        const schedule = {};
        letters.forEach((letter) => {
          schedule[letter] = [];
        });

        let current = circle.slice();
        for (let r = 0; r < rounds; r += 1) {
          for (let i = 0; i < letters.length / 2; i += 1) {
            const left = current[i];
            const right = current[current.length - 1 - i];
            schedule[left].push(right);
            schedule[right].push(left);
          }
          const fixed = current[0];
          const rotating = current.slice(1);
          const last = rotating.pop();
          current = [fixed, last, ...rotating];
        }

        return schedule;
      };

      const SCHEDULE_BY_LETTER = buildSchedule(LETTERS);

      const buildRoundMap = (letters) => {
        const schedule = buildSchedule(letters);
        const roundMap = new Map();
        letters.forEach((letter) => {
          schedule[letter].forEach((opponent, index) => {
            const key = letter < opponent ? `${letter}-${opponent}` : `${opponent}-${letter}`;
            if (!roundMap.has(key)) roundMap.set(key, index + 1);
          });
        });
        return roundMap;
      };

      const ROUND_MAP = buildRoundMap(LETTERS);

      const getMatchLettersFromKey = (key) => {
        if (!key) return null;
        const match = String(key).trim().match(/^([A-L])-([A-L])$/);
        if (!match) return null;
        return { rowLetter: match[1], colLetter: match[2] };
      };

      const getRoundNumberFromKey = (key) => {
        const match = getMatchLettersFromKey(key);
        if (!match) return null;
        const normalized =
          match.rowLetter < match.colLetter
            ? `${match.rowLetter}-${match.colLetter}`
            : `${match.colLetter}-${match.rowLetter}`;
        return ROUND_MAP.get(normalized) ?? null;
      };

      const parseCsv = (text) => {
        const rows = [];
        let row = [];
        let field = "";
        let inQuotes = false;

        for (let i = 0; i < text.length; i += 1) {
          const char = text[i];
          const next = text[i + 1];

          if (char === '"') {
            if (inQuotes && next === '"') {
              field += '"';
              i += 1;
            } else {
              inQuotes = !inQuotes;
            }
            continue;
          }

          if (!inQuotes && (char === "," || char === "\n" || char === "\r")) {
            if (char === ",") {
              row.push(field);
              field = "";
              continue;
            }

            if (char === "\r" && next === "\n") {
              i += 1;
            }

            row.push(field);
            field = "";
            if (row.some((value) => value.trim() !== "")) {
              rows.push(row);
            }
            row = [];
            continue;
          }

          field += char;
        }

        row.push(field);
        if (row.some((value) => value.trim() !== "")) {
          rows.push(row);
        }
        return rows;
      };

      const rowsToObjects = (rows) => {
        if (!rows.length) return [];
        const headers = rows[0];
        return rows.slice(1).map((row) => {
          const obj = {};
          headers.forEach((header, idx) => {
            obj[header] = row[idx] ?? "";
          });
          return obj;
        });
      };

      const groupByPlayer = (items) => {
        const map = new Map();
        items.forEach((item) => {
          const name = item.player || "(不明)";
          if (!map.has(name)) map.set(name, []);
          map.get(name).push(item);
        });
        return map;
      };

      const normalizeLetter = (value) => {
        if (!value) return "";
        const text = String(value).trim().toUpperCase();
        return LETTERS.includes(text) ? text : "";
      };

      const buildLetterMaps = (items) => {
        const letterToName = new Map();
        const nameToLetter = new Map();
        items.forEach((item) => {
          const letter = normalizeLetter(item.letter);
          const name = item.player || "(不明)";
          if (!letter) return;
          if (!letterToName.has(letter)) letterToName.set(letter, name);
          if (!nameToLetter.has(name)) nameToLetter.set(name, letter);
        });
        return { letterToName, nameToLetter };
      };

      const buildMatchKeys = (row) => {
        const linkKeys = Object.keys(row).filter((key) => key.startsWith("match"));
        return linkKeys.map((key) => row[key]).filter(Boolean);
      };

      const resolveManifestHref = (manifestPath, href) => {
        if (!manifestPath || !href) return "";
        try {
          const manifestUrl = new URL(manifestPath, location.href);
          return new URL(href, manifestUrl).toString();
        } catch (err) {
          return href;
        }
      };

      const loadManifest = async (manifestPath) => {
        if (!manifestPath) return null;
        try {
          const response = await fetch(manifestPath, { cache: "no-store" });
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          return await response.json();
        } catch (err) {
          console.warn("manifest.jsonの読み込みに失敗しました。", err);
          return null;
        }
      };

      const extractMonsterSlug = (imageUrl) => {
        if (!imageUrl) return "";
        const clean = String(imageUrl).split("?")[0];
        const base = clean.split("/").pop() || "";
        const namePart = base.replace(/\.[^.]+$/, "");
        let parts = namePart.split("_");
        if (parts.length && /^\d+$/.test(parts[0])) {
          parts = parts.slice(1);
        }
        if (parts.length && /^\d+g\d+$/i.test(parts[parts.length - 1])) {
          parts = parts.slice(0, -1);
        }
        const slug = parts.join("_").trim();
        return slug || "";
      };

      const buildSkillList = (text) => {
        if (!text) return [];
        return text
          .split("/")
          .map((item) => item.trim())
          .filter(Boolean);
      };

      const setSearch = (value) => {
        currentQuery = value;
        searchInput.value = value;
        if (currentPhase) renderPhase(currentPhase);
      };

      const normalizeSearchText = (value) => {
        return String(value || "")
          .normalize("NFKC")
          .toLowerCase()
          .replace(/[\u30a1-\u30f6]/g, (char) =>
            String.fromCharCode(char.charCodeAt(0) - 0x60)
          );
      };

      const rowMatchesQuery = (row, query) => {
        if (!query) return true;
        const q = normalizeSearchText(query);
        const fields = [
          row.monster,
          row.image,
          row["A(アクティブ)"],
          row["P(コンパニオン)"],
        ];
        return fields.some((value) => normalizeSearchText(value).includes(q));
      };

      const buildGroupCard = async (groupLabel, groupConfig, query) => {
        const card = document.createElement("article");
        card.className = "group-card";
        let hitCount = 0;
        const csvPath = groupConfig?.csv || "";
        const manifestPath = groupConfig?.manifest || "";

        const header = document.createElement("div");
        header.className = "group-header";

        const title = document.createElement("h3");
        title.className = "group-title";
        title.textContent = groupLabel;

        const actions = document.createElement("div");
        actions.className = "group-actions";

        const download = document.createElement("a");
        download.className = "download-button";
        download.textContent = "CSVダウンロード";
        if (csvPath) {
          download.href = csvPath;
          download.setAttribute("download", "");
        } else {
          download.classList.add("is-disabled");
          download.href = "#";
        }

        actions.append(download);
        header.append(title, actions);
        card.appendChild(header);

        if (!csvPath) {
          const empty = document.createElement("div");
          empty.className = "empty";
          empty.textContent = "CSVが設定されていません。";
          card.appendChild(empty);
          return { card, hitCount };
        }

        try {
          const response = await fetch(csvPath);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const manifest = await loadManifest(manifestPath);
          const manifestMatches = manifest?.matches || {};
          const text = await response.text();
          const rows = parseCsv(text);
          const data = rowsToObjects(rows);
          const filtered = query ? data.filter((row) => rowMatchesQuery(row, query)) : data;
          hitCount = query ? filtered.length : 0;
          const { letterToName, nameToLetter } = buildLetterMaps(data);
          const hasLetters = letterToName.size > 0;

          if (!data.length) {
            const empty = document.createElement("div");
            empty.className = "empty";
            empty.textContent = "CSVの中身が空です。";
            card.appendChild(empty);
            return { card, hitCount };
          }

          if (!filtered.length) {
            const empty = document.createElement("div");
            empty.className = "empty";
            empty.textContent = "検索条件に一致するデータがありません。";
            card.appendChild(empty);
            return { card, hitCount };
          }

          const players = groupByPlayer(filtered);
          const list = document.createElement("div");
          list.className = "player-list";

          players.forEach((rows, playerName) => {
            const playerCard = document.createElement("details");
            playerCard.className = "player-card player-accordion";
            playerCard.open = true;

            const summary = document.createElement("summary");
            summary.className = "player-summary";

            const nameEl = document.createElement("h4");
            nameEl.className = "player-name";
            nameEl.textContent = playerName;

            summary.append(nameEl);
            playerCard.appendChild(summary);

            const playerLetter = hasLetters ? nameToLetter.get(playerName) || "" : "";
            const schedule = playerLetter ? SCHEDULE_BY_LETTER[playerLetter] || [] : [];

            const table = document.createElement("table");
            const thead = document.createElement("thead");
            thead.innerHTML =
              "<tr>" +
              "<th>名称</th>" +
              "<th class=\"lv-col\">Lv</th>" +
              "<th class=\"variant-cell variant-col\">種別</th>" +
              "<th class=\"lv-variant-combo\"><span>Lv</span><span>／</span><span>種</span></th>" +
              "<th class=\"image-cell\">画像</th>" +
              "<th class=\"skill-cell\">A</th>" +
              "<th class=\"skill-cell\">P</th>" +
              "<th class=\"is-toggle\" data-role=\"appear-toggle\">出場数</th>" +
              "</tr>";
            table.appendChild(thead);

            const tbody = document.createElement("tbody");
            rows
              .slice()
              .sort((a, b) => Number(b["出場回数"] || 0) - Number(a["出場回数"] || 0))
              .forEach((row) => {
              const tr = document.createElement("tr");

              const monster = document.createElement("td");
              const monsterSlug = extractMonsterSlug(row.image) || "";
              const monsterLabel = row.monster || "";
              const monsterButton = document.createElement("button");
              monsterButton.type = "button";
              monsterButton.className = "search-link";
              monsterButton.textContent = monsterLabel;
              monsterButton.addEventListener("click", () => {
                if (monsterSlug) {
                  setSearch(monsterSlug);
                } else if (monsterLabel) {
                  setSearch(monsterLabel);
                }
              });
              monster.appendChild(monsterButton);

              const level = document.createElement("td");
              level.className = "lv-col";
              level.textContent = row.level || "";

              const variant = document.createElement("td");
              variant.className = "variant-cell variant-col";
              variant.textContent = row.variant || "";

              const lvVariant = document.createElement("td");
              lvVariant.className = "lv-variant-combo";
              const lvSpan = document.createElement("span");
              lvSpan.textContent = row.level || "";
              const slashSpan = document.createElement("span");
              slashSpan.textContent = row.level && row.variant ? "／" : "";
              const variantSpan = document.createElement("span");
              variantSpan.textContent = row.variant || "";
              lvVariant.append(lvSpan, slashSpan, variantSpan);

              const imageCell = document.createElement("td");
              imageCell.className = "image-cell";
              if (row.image) {
                const imgButton = document.createElement("button");
                imgButton.type = "button";
                imgButton.className = "image-button";
                imgButton.addEventListener("click", () => {
                  if (monsterSlug) {
                    setSearch(monsterSlug);
                  }
                });
                const img = document.createElement("img");
                img.src = row.image;
                img.alt = row.monster || "";
                imgButton.appendChild(img);
                imageCell.appendChild(imgButton);
              }

              const action = document.createElement("td");
              const actionSkills = buildSkillList(row["A(アクティブ)"]);
              if (actionSkills.length) {
                const list = document.createElement("div");
                list.className = "skill-list";
                actionSkills.forEach((skill) => {
                  const tag = document.createElement("button");
                  tag.type = "button";
                  tag.className = "skill-tag";
                  tag.textContent = skill;
                  tag.addEventListener("click", () => setSearch(skill));
                  list.appendChild(tag);
                });
                action.appendChild(list);
              } else {
                action.textContent = row["A(アクティブ)"] || "";
              }

              const p0 = document.createElement("td");
              const p0Skills = buildSkillList(row["P(コンパニオン)"]);
              if (p0Skills.length) {
                const list = document.createElement("div");
                list.className = "skill-list";
                p0Skills.forEach((skill) => {
                  const tag = document.createElement("button");
                  tag.type = "button";
                  tag.className = "skill-tag";
                  tag.textContent = skill;
                  tag.addEventListener("click", () => setSearch(skill));
                  list.appendChild(tag);
                });
                p0.appendChild(list);
              } else {
                p0.textContent = row["P(コンパニオン)"] || "";
              }

              const appear = document.createElement("td");
              appear.className = "appear-cell";
              const appearCount = row["出場回数"] || "";
              const matchKeys = buildMatchKeys(row);
              if (matchKeys.length) {
                const details = document.createElement("details");
                details.className = "appear-toggle";
                const summary = document.createElement("summary");
                summary.textContent = appearCount || String(matchKeys.length);
                const linkList = document.createElement("div");
                linkList.className = "link-list";
                matchKeys.forEach((matchKey, index) => {
                  const entry = manifestMatches[matchKey] || null;
                  const match = entry
                    ? { rowLetter: entry.leftLetter, colLetter: entry.rightLetter }
                    : getMatchLettersFromKey(matchKey);
                  const roundNumber = entry?.round || getRoundNumberFromKey(matchKey);
                  let label = roundNumber ? `${roundNumber}戦目` : `対戦${index + 1}`;
                  if (
                    hasLetters &&
                    match &&
                    playerLetter &&
                    (match.rowLetter === playerLetter || match.colLetter === playerLetter)
                  ) {
                    const opponentLetter =
                      match.rowLetter === playerLetter ? match.colLetter : match.rowLetter;
                    const opponentName = letterToName.get(opponentLetter) || opponentLetter;
                    const roundIndex = schedule.indexOf(opponentLetter);
                    if (roundIndex >= 0 && opponentName) {
                      label = `${roundIndex + 1}.${opponentName}戦`;
                    }
                  }
                  const href = entry?.html ? resolveManifestHref(manifestPath, entry.html) : "";
                  if (href) {
                    const link = document.createElement("a");
                    link.href = href;
                    link.target = "_blank";
                    link.rel = "noopener";
                    link.textContent = label;
                    linkList.appendChild(link);
                  } else {
                    const text = document.createElement("span");
                    text.textContent = label;
                    linkList.appendChild(text);
                  }
                });
                details.append(summary, linkList);
                appear.appendChild(details);
              } else {
                appear.textContent = appearCount;
              }

              tr.append(monster, level, variant, lvVariant, imageCell, action, p0, appear);
              tbody.appendChild(tr);
            });

            table.appendChild(tbody);

            const toggleHeader = table.querySelector("th[data-role=\"appear-toggle\"]");
            if (toggleHeader) {
              toggleHeader.addEventListener("click", () => {
                const detailsList = table.querySelectorAll("details.appear-toggle");
                if (!detailsList.length) return;
                const shouldOpen = Array.from(detailsList).some((detail) => !detail.open);
                detailsList.forEach((detail) => {
                  detail.open = shouldOpen;
                });
              });
            }
            const body = document.createElement("div");
            body.className = "player-body";
            body.appendChild(table);
            playerCard.appendChild(body);
            list.appendChild(playerCard);
          });

          card.appendChild(list);
        } catch (err) {
          const empty = document.createElement("div");
          empty.className = "empty";
          empty.textContent = "CSVが見つかりません。ファイル名とパスを確認してください。";
          card.appendChild(empty);
        }

        return { card, hitCount };
      };

      const renderPhase = async (phase) => {
        if (!phase || !phase.groups) {
          renderMessage("期の設定が不正です。phase_map.json を確認してください。");
          updateSearchStatus(0);
          return;
        }
        groupGrid.innerHTML = "";
        let totalHits = 0;
        for (const group of GROUP_ORDER) {
          const { card, hitCount } = await buildGroupCard(group, phase.groups[group], currentQuery);
          totalHits += hitCount;
          groupGrid.appendChild(card);
        }
        updateSearchStatus(totalHits);
      };

      const setupTabs = () => {
        phaseTabs.innerHTML = "";
        if (!PHASES.length) {
          renderMessage("期の設定がありません。phase_map.json を確認してください。");
          return;
        }
        phaseMeta = buildPhaseMeta(PHASES);
        phaseByNumber = new Map();
        phaseMeta.forEach((meta) => {
          if (meta.number != null) phaseByNumber.set(meta.number, meta);
        });
        const latestTwo = phaseMeta
          .slice()
          .sort((a, b) => b.number - a.number || b.index - a.index)
          .slice(0, 2);

        latestTwo.forEach((meta) => {
          const button = document.createElement("button");
          button.type = "button";
          button.className = "phase-button";
          button.textContent = meta.phase.label;
          button.dataset.index = String(meta.index);
          button.addEventListener("click", () => {
            setActivePhase(meta);
          });
          phaseTabs.appendChild(button);
        });

        const latestMeta = latestTwo[0] || phaseMeta[0];
        if (latestMeta) setActivePhase(latestMeta);
      };

      searchInput.addEventListener("input", (event) => {
        currentQuery = event.target.value.trim();
        if (currentPhase) renderPhase(currentPhase);
      });

      if (openAllPlayersButton) {
        openAllPlayersButton.addEventListener("click", () => {
          setAllPlayerAccordions(true);
        });
      }

      if (closeAllPlayersButton) {
        closeAllPlayersButton.addEventListener("click", () => {
          setAllPlayerAccordions(false);
        });
      }

      if (phaseJumpForm) {
        phaseJumpForm.addEventListener("submit", (event) => {
          event.preventDefault();
          const value = parseInt(phaseJumpInput?.value || "", 10);
          if (Number.isNaN(value)) {
            return;
          }
          const meta = phaseByNumber.get(value);
          if (!meta) {
            return;
          }
          setActivePhase(meta);
        });
      }

      const init = async () => {
        PHASES = await loadPhaseMap();
        setupTabs();
      };

      init();
