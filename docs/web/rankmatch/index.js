let PHASES = [];

      const GROUP_ORDER = ["S", "A1", "A2"];

      const phaseTabs = document.getElementById("phaseTabs");
      const groupGrid = document.getElementById("groupGrid");
      const searchInput = document.getElementById("searchInput");
      const searchStatus = document.getElementById("searchStatus");
      const rankingSummary = document.getElementById("rankingSummary");
      const openAllPlayersButton = document.getElementById("openAllPlayers");
      const closeAllPlayersButton = document.getElementById("closeAllPlayers");
      const phaseJumpForm = document.getElementById("phaseJumpForm");
      const phaseJumpInput = document.getElementById("phaseJumpInput");
      let currentPhase = null;
      let currentQuery = "";
      let phaseMeta = [];
      let phaseByNumber = new Map();
      let renderSequence = 0;

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
          const numbers = Array.isArray(data) ? data : data.phases;
          return Array.isArray(numbers)
            ? numbers.map((number) => Number(number)).filter(Number.isFinite)
            : [];
        } catch (err) {
          console.warn("phase_map.jsonの読み込みに失敗しました。", err);
          return [];
        }
      };

      const buildGroupConfig = (phaseNumber, group) => ({
        csv: `./data/${phaseNumber}/${group}/skill_list.csv`,
        manifest: `./data/${phaseNumber}/${group}/manifest.json`,
      });

      const buildPhaseMeta = (phases) => {
        return phases.map((number, index) => ({
          phase: {
            number,
            groups: Object.fromEntries(
              GROUP_ORDER.map((group) => [group, buildGroupConfig(number, group)])
            ),
          },
          index,
          number,
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

      const addRankingCount = (counts, name, count) => {
        if (!name || count <= 0) return;
        counts.set(name, (counts.get(name) || 0) + count);
      };

      const buildRankingMonsterImage = (imageUrl) => {
        if (!imageUrl) return "";
        return String(imageUrl).replace(/_\d+g\d+(?=\.[^./?#]+(?:[?#]|$))/i, "_2g7");
      };

      const isTargetMonsterImage = (imageUrl) => {
        return /_2g7(?=\.[^./?#]+(?:[?#]|$))/i.test(String(imageUrl || ""));
      };

      const addRankingImageCandidate = (entry, imageUrl, prefer = false) => {
        if (!imageUrl || entry.images.includes(imageUrl)) return;
        if (prefer) {
          entry.images.unshift(imageUrl);
        } else {
          entry.images.push(imageUrl);
        }
        entry.image = entry.images[0] || "";
      };

      const addMonsterRankingCount = (counts, row, count) => {
        const name = extractMonsterSlug(row.image);
        if (!name || count <= 0) return;
        const image = buildRankingMonsterImage(row.image);
        const current = counts.get(name);
        if (current) {
          current.count += count;
          addRankingImageCandidate(current, image, isTargetMonsterImage(row.image));
          return;
        }
        const entry = {
          name,
          count,
          image: "",
          images: [],
        };
        addRankingImageCandidate(entry, image, isTargetMonsterImage(row.image));
        counts.set(name, entry);
      };

      const sortRankingEntries = (entries) => {
        return entries.sort(
          (a, b) => b.count - a.count || a.name.localeCompare(b.name, "ja")
        );
      };

      const buildRanking = (rows, field) => {
        const counts = new Map();
        rows.forEach((row) => {
          const appearanceCount = Number(row["出場回数"] || 0);
          if (field === "monster") {
            addMonsterRankingCount(counts, row, appearanceCount);
            return;
          }
          buildSkillList(row[field]).forEach((skill) => {
            addRankingCount(counts, skill, appearanceCount);
          });
        });
        const entries =
          field === "monster"
            ? Array.from(counts.values())
            : Array.from(counts, ([name, count]) => ({ name, count }));
        return sortRankingEntries(entries).slice(0, 10);
      };

      const buildPointRanking = (rows, field) => {
        const counts = new Map();
        rows.forEach((row) => {
          if (field === "monster") {
            addMonsterRankingCount(counts, row, 1);
            return;
          }
          new Set(buildSkillList(row[field])).forEach((skill) => {
            addRankingCount(counts, skill, 1);
          });
        });
        const entries =
          field === "monster"
            ? Array.from(counts.values())
            : Array.from(counts, ([name, count]) => ({ name, count }));
        return sortRankingEntries(entries);
      };

      const PIE_COLORS = ["#6750a4", "#386a20", "#984061", "#00639b", "#7d5700"];
      const PIE_OTHER_COLOR = "#cac4d0";

      const buildRankingImageCandidates = (entry) => {
        const candidates = [];
        const addCandidate = (url) => {
          if (url && !candidates.includes(url)) candidates.push(url);
        };
        (entry.images || [entry.image]).forEach(addCandidate);
        candidates.slice().forEach((url) => {
          ["g", "p", "s", "b"].forEach((directory) => {
            addCandidate(url.replace(/\/mi\/[bgps]\//i, `/mi/${directory}/`));
          });
        });
        return candidates;
      };

      const createRankingName = (entry) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "search-link ranking-search ranking-name";
        button.addEventListener("click", () => setSearch(entry.name));
        if (entry.image) {
          const imageCandidates = buildRankingImageCandidates(entry);
          let imageIndex = 0;
          const image = document.createElement("img");
          image.className = "ranking-monster-image";
          image.src = imageCandidates[imageIndex];
          image.alt = entry.name;
          image.loading = "lazy";
          image.addEventListener("error", () => {
            imageIndex += 1;
            if (imageIndex < imageCandidates.length) {
              image.src = imageCandidates[imageIndex];
            } else {
              image.hidden = true;
            }
          });
          button.appendChild(image);
        }
        const label = document.createElement("span");
        label.textContent = entry.name;
        button.appendChild(label);
        return button;
      };

      const appendChartRanking = (card, entries) => {
        const topEntries = entries.slice(0, 5);
        const total = entries.reduce((sum, entry) => sum + entry.count, 0);
        if (!total) return;

        const otherCount = entries.slice(5).reduce((sum, entry) => sum + entry.count, 0);
        const slices = topEntries.map((entry, index) => ({
          ...entry,
          color: PIE_COLORS[index],
        }));
        if (otherCount) {
          slices.push({ name: "その他", count: otherCount, color: PIE_OTHER_COLOR });
        }

        let currentPercent = 0;
        const gradientParts = slices.map((slice) => {
          const start = currentPercent;
          currentPercent += (slice.count / total) * 100;
          return `${slice.color} ${start}% ${currentPercent}%`;
        });

        const layout = document.createElement("div");
        layout.className = "ranking-chart-layout";
        const pie = document.createElement("div");
        pie.className = "ranking-pie";
        pie.style.background = `conic-gradient(${gradientParts.join(", ")})`;
        pie.setAttribute("role", "img");
        pie.setAttribute(
          "aria-label",
          slices
            .map((slice) => `${slice.name} ${((slice.count / total) * 100).toFixed(1)}%`)
            .join("、")
        );

        const list = document.createElement("ol");
        list.className = "ranking-list ranking-chart-list";
        topEntries.forEach((entry, index) => {
          const item = document.createElement("li");
          item.className = "ranking-item";
          const content = document.createElement("div");
          content.className = "ranking-chart-entry";
          const color = document.createElement("span");
          color.className = "ranking-color";
          color.style.background = PIE_COLORS[index];
          const name = createRankingName(entry);
          const count = document.createElement("span");
          count.className = "ranking-count";
          count.textContent = `${entry.count}点・${((entry.count / total) * 100).toFixed(1)}%`;
          content.append(color, name, count);
          item.appendChild(content);
          list.appendChild(item);
        });
        layout.append(pie, list);
        card.appendChild(layout);

        if (otherCount) {
          const other = document.createElement("div");
          other.className = "ranking-other";
          const color = document.createElement("span");
          color.className = "ranking-color";
          color.style.background = PIE_OTHER_COLOR;
          const name = document.createElement("span");
          name.textContent = "その他";
          const count = document.createElement("span");
          count.className = "ranking-count";
          count.textContent = `${otherCount}点・${((otherCount / total) * 100).toFixed(1)}%`;
          other.append(color, name, count);
          card.appendChild(other);
        }
      };

      const getSearchAggregation = (rows, query) => {
        const q = normalizeSearchText(query);
        const monsterRows = rows.filter((row) => {
          const values = [extractMonsterSlug(row.image), row.monster, row.image];
          return values.some((value) => normalizeSearchText(value).includes(q));
        });
        const skillRows = rows.filter((row) => {
          const skills = [
            ...buildSkillList(row["A(アクティブ)"]),
            ...buildSkillList(row["P(コンパニオン)"]),
          ];
          return skills.some((skill) => normalizeSearchText(skill).includes(q));
        });
        const exactMonster = monsterRows.some(
          (row) =>
            normalizeSearchText(extractMonsterSlug(row.image)) === q ||
            normalizeSearchText(row.monster) === q
        );
        const exactSkill = skillRows.some((row) =>
          [
            ...buildSkillList(row["A(アクティブ)"]),
            ...buildSkillList(row["P(コンパニオン)"]),
          ].some((skill) => normalizeSearchText(skill) === q)
        );

        if (exactSkill && !exactMonster) return { kind: "skill", rows: skillRows };
        if (exactMonster) return { kind: "monster", rows: monsterRows };
        if (!monsterRows.length && skillRows.length) return { kind: "skill", rows: skillRows };
        if (monsterRows.length && !skillRows.length) return { kind: "monster", rows: monsterRows };
        if (skillRows.length >= monsterRows.length) return { kind: "skill", rows: skillRows };
        if (monsterRows.length) return { kind: "monster", rows: monsterRows };
        return { kind: "", rows: [] };
      };

      const appendChartCard = (titleText, entries) => {
        if (!entries.length) return;
        const card = document.createElement("article");
        card.className = "ranking-card has-chart";
        const title = document.createElement("h2");
        title.className = "ranking-title";
        title.textContent = titleText;
        card.appendChild(title);
        appendChartRanking(card, entries);
        rankingSummary.appendChild(card);
      };

      const renderRankingSummary = (rows) => {
        if (!rankingSummary) return;
        rankingSummary.innerHTML = "";
        rankingSummary.classList.toggle("is-search-result", Boolean(currentQuery));
        if (!rows.length) {
          rankingSummary.hidden = true;
          return;
        }

        if (currentQuery) {
          const aggregation = getSearchAggregation(rows, currentQuery);
          if (aggregation.kind === "monster") {
            appendChartCard(
              `「${currentQuery}」のAスキル トップ5`,
              buildPointRanking(aggregation.rows, "A(アクティブ)")
            );
            appendChartCard(
              `「${currentQuery}」のPスキル トップ5`,
              buildPointRanking(aggregation.rows, "P(コンパニオン)")
            );
          } else if (aggregation.kind === "skill") {
            appendChartCard(
              `「${currentQuery}」の採用モンスター トップ5`,
              buildPointRanking(aggregation.rows, "monster")
            );
          }
          rankingSummary.hidden = !rankingSummary.children.length;
          return;
        }

        const rankings = [
          ["採用モンスター", buildRanking(rows, "monster")],
          ["Aスキル", buildRanking(rows, "A(アクティブ)")],
          ["Pスキル", buildRanking(rows, "P(コンパニオン)")],
        ];

        rankings.forEach(([titleText, entries]) => {
          const card = document.createElement("article");
          card.className = "ranking-card";
          const title = document.createElement("h2");
          title.className = "ranking-title";
          title.textContent = `${titleText} トップ10`;
          const list = document.createElement("ol");
          list.className = "ranking-list";

          entries.forEach((entryData) => {
            const { count } = entryData;
            const item = document.createElement("li");
            item.className = "ranking-item";
            const entry = document.createElement("div");
            entry.className = "ranking-entry";
            const name = createRankingName(entryData);
            const countText = document.createElement("span");
            countText.className = "ranking-count";
            countText.textContent = `${count}回`;
            entry.append(name, countText);
            item.appendChild(entry);
            list.appendChild(item);
          });

          card.append(title, list);
          rankingSummary.appendChild(card);
        });
        rankingSummary.hidden = false;
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
        let data = [];
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
          return { card, hitCount, data };
        }

        try {
          const response = await fetch(csvPath);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const manifest = await loadManifest(manifestPath);
          const manifestMatches = manifest?.matches || {};
          const text = await response.text();
          const rows = parseCsv(text);
          data = rowsToObjects(rows);
          const filtered = query ? data.filter((row) => rowMatchesQuery(row, query)) : data;
          hitCount = query ? filtered.length : 0;
          const { letterToName, nameToLetter } = buildLetterMaps(data);
          const hasLetters = letterToName.size > 0;

          if (!data.length) {
            const empty = document.createElement("div");
            empty.className = "empty";
            empty.textContent = "CSVの中身が空です。";
            card.appendChild(empty);
            return { card, hitCount, data };
          }

          if (!filtered.length) {
            const empty = document.createElement("div");
            empty.className = "empty";
            empty.textContent = "検索条件に一致するデータがありません。";
            card.appendChild(empty);
            return { card, hitCount, data };
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

        return { card, hitCount, data };
      };

      const renderPhase = async (phase) => {
        const renderId = ++renderSequence;
        if (!phase || !phase.groups) {
          renderMessage("期の設定が不正です。phase_map.json を確認してください。");
          updateSearchStatus(0);
          renderRankingSummary([]);
          return;
        }
        groupGrid.innerHTML = "";
        if (rankingSummary) rankingSummary.hidden = true;
        const query = currentQuery;
        const results = await Promise.all(
          GROUP_ORDER.map((group) => buildGroupCard(group, phase.groups[group], query))
        );
        if (renderId !== renderSequence) return;
        const totalHits = results.reduce((total, result) => total + result.hitCount, 0);
        const allRows = results.flatMap((result) => result.data);
        results.forEach(({ card }) => groupGrid.appendChild(card));
        updateSearchStatus(totalHits);
        renderRankingSummary(allRows);
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
          button.textContent = `${meta.number}期`;
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
