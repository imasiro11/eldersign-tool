(() => {
  const common = window.EldersignToolCommon || {};
  const renderResultPairs = common.renderResultPairs || ((container, pairs) => {
    if (!container) return;
    container.innerHTML = "";
    (pairs || []).forEach((pair) => {
      const item = document.createElement("div");
      item.className = "result-row";
      const label = document.createElement("div");
      label.className = "result-label";
      label.textContent = pair && pair.label != null ? String(pair.label) : "";
      const value = document.createElement("div");
      value.className = "result-value";
      value.textContent = pair && pair.value != null ? String(pair.value) : "-";
      item.append(label, value);
      container.appendChild(item);
    });
  });
  const setChipValue = common.setChipValue || ((input, value) => {
    if (!input || value == null) return;
    const stringValue = String(value);
    input.value = stringValue;
    const group = document.querySelector(`.chip-group[data-chip-target="${input.id}"]`);
    if (!group) return;
    group.querySelectorAll(".chip-button").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.value === stringValue);
    });
  });
  const bindChipGroups = common.bindChipGroups || ((onChange) => {
    document.querySelectorAll(".chip-group").forEach((group) => {
      group.addEventListener("click", (event) => {
        const button = event.target.closest(".chip-button");
        if (!button || !group.contains(button)) return;
        const targetId = group.dataset.chipTarget;
        if (!targetId) return;
        const input = document.getElementById(targetId);
        if (!input) return;
        setChipValue(input, button.dataset.value);
        if (typeof onChange === "function") onChange(input, button.dataset.value, group);
      });
    });
  });

  const inputs = {
    level: document.getElementById("level"),
    family: document.getElementById("family"),
    grade: document.getElementById("grade"),
    rarity: document.getElementById("rarity"),
    skillInherit: document.getElementById("skill-inherit"),
    abilityNourishment: document.getElementById("ability-nourishment"),
  };
  const outputList = document.getElementById("exp-result-list");
  const quickTableBody = document.getElementById("growth-quick-table-body");
  const levelPresetButtons = document.getElementById("level-preset-buttons");
  const STORAGE_KEY = "eldersign_exp_form_v1";

  const RARITY_CONFIG = {
    1: { label: "ブロンズ", maxLevel: 30, growBase: 20, growStep: 10 },
    2: { label: "シルバー", maxLevel: 50, growBase: 25, growStep: 15 },
    4: { label: "ゴールド", maxLevel: 70, growBase: 30, growStep: 20 },
    8: { label: "プラチナ", maxLevel: 90, growBase: 35, growStep: 25 },
  };
  const QUICK_LEVELS = [30, 50, 70, 90];

  function readNumber(input, fallback = 0) {
    const value = Number(input && input.value);
    return Number.isFinite(value) ? value : fallback;
  }

  function calcGrowthExp(config, targetLevel) {
    if (!config) return 0;
    const level = Math.max(1, Math.min(config.maxLevel, Math.floor(targetLevel)));
    const terms = Math.max(0, level - 1);
    return (terms * (2 * config.growBase + (terms - 1) * config.growStep)) / 2;
  }

  function syncLevelPresetActive(level) {
    if (!levelPresetButtons) return;
    levelPresetButtons.querySelectorAll(".chip-button").forEach((button) => {
      button.classList.toggle("is-active", Number(button.dataset.level) === Number(level));
    });
  }

  function renderQuickTable() {
    if (!quickTableBody) return;
    quickTableBody.innerHTML = "";

    Object.keys(RARITY_CONFIG).forEach((key) => {
      const config = RARITY_CONFIG[key];
      const tr = document.createElement("tr");

      const labelCell = document.createElement("th");
      labelCell.textContent = `${config.label} (MAX Lv${config.maxLevel})`;
      tr.appendChild(labelCell);

      QUICK_LEVELS.forEach((level) => {
        const td = document.createElement("td");
        if (level > config.maxLevel) {
          td.textContent = "-";
        } else {
          const exp = calcGrowthExp(config, level);
          td.textContent = `${exp.toLocaleString("ja-JP")}`;
        }
        tr.appendChild(td);
      });

      quickTableBody.appendChild(tr);
    });
  }

  function saveFormState() {
    const state = {};
    Object.entries(inputs).forEach(([key, input]) => {
      if (!input) return;
      state[key] = input.type === "checkbox" ? input.checked : input.value;
    });
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (error) {
      // Ignore storage failures (private mode, quota, etc.)
    }
  }

  function loadFormState() {
    let state = null;
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      state = raw ? JSON.parse(raw) : null;
    } catch (error) {
      state = null;
    }
    if (!state || typeof state !== "object") return;

    Object.entries(inputs).forEach(([key, input]) => {
      if (!input || !(key in state)) return;
      if (input.type === "checkbox") {
        input.checked = Boolean(state[key]);
      } else {
        input.value = String(state[key]);
      }
    });
  }

  function calculate() {
    const selectedRarity = readNumber(inputs.rarity, 1);
    const rarityConfig = RARITY_CONFIG[selectedRarity] || RARITY_CONFIG[1];
    const maxLevel = rarityConfig.maxLevel;
    const level = Math.max(1, Math.min(maxLevel, Math.floor(readNumber(inputs.level, 1))));
    const family = readNumber(inputs.family, 1);
    const grade = readNumber(inputs.grade, 1);
    const rarity = selectedRarity;
    const skillValue = inputs.skillInherit.checked ? 0.5 : 0;
    const nourishmentValue = inputs.abilityNourishment.checked ? 10 : 1;

    const base = 16 + 3.2 * (level - 1);
    const factor = skillValue + nourishmentValue;
    const raw = base * family * grade * rarity * factor;
    const total = Math.floor(raw);
    const growthExp = calcGrowthExp(rarityConfig, level);
    const diff = total - growthExp;
    const diffText = `${diff >= 0 ? "+" : ""}${diff.toLocaleString("ja-JP")}`;

    renderResultPairs(
      outputList,
      [
        { label: "合成経験値", value: `${total.toLocaleString("ja-JP")}` },
        { label: `育成必要Exp (Lv1→Lv${level})`, value: `${growthExp.toLocaleString("ja-JP")}` },
        { label: "差分 (合成経験値 - 育成必要Exp)", value: diffText },
      ],
      { itemClass: "result-row" }
    );

    if (inputs.level.value !== String(level)) {
      inputs.level.value = String(level);
    }
    inputs.level.max = String(maxLevel);
    syncLevelPresetActive(level);
    saveFormState();
  }

  Object.values(inputs).forEach((input) => {
    if (!input) return;
    input.addEventListener(input.type === "checkbox" ? "change" : "input", calculate);
  });

  if (levelPresetButtons) {
    levelPresetButtons.addEventListener("click", (event) => {
      const button = event.target.closest(".chip-button");
      if (!button || !levelPresetButtons.contains(button)) return;
      const level = Number(button.dataset.level);
      if (!Number.isFinite(level)) return;
      inputs.level.value = String(level);
      calculate();
    });
  }

  loadFormState();
  setChipValue(inputs.family, inputs.family.value);
  setChipValue(inputs.rarity, inputs.rarity.value);
  bindChipGroups(() => {
    calculate();
  });
  renderQuickTable();
  calculate();
})();
