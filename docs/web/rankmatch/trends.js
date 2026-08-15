const GROUPS = ["S", "A1", "A2"];
const TYPES = {
  monster: { field: "monster", label: "モンスター" },
  action: { field: "A(アクティブ)", label: "Aスキル" },
  passive: { field: "P(コンパニオン)", label: "Pスキル" },
};
const SERIES_COLORS = ["#6750a4", "#2e7d32", "#c23b6e", "#0277bd", "#b26a00"];
const SVG_NS = "http://www.w3.org/2000/svg";

const typeTabs = document.getElementById("typeTabs");
const itemPicker = document.getElementById("itemPicker");
const itemInput = document.getElementById("itemInput");
const itemOptions = document.getElementById("itemOptions");
const selectedItems = document.getElementById("selectedItems");
const loadingStatus = document.getElementById("loadingStatus");
const chartWrap = document.getElementById("chartWrap");
const trendChart = document.getElementById("trendChart");
const chartNote = document.getElementById("chartNote");

let phaseData = [];
let currentType = "monster";
let selectedNames = [];

const parseCsv = (text) => {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        index += 1;
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
      if (char === "\r" && next === "\n") index += 1;
      row.push(field);
      field = "";
      if (row.some((value) => value.trim())) rows.push(row);
      row = [];
      continue;
    }
    field += char;
  }
  row.push(field);
  if (row.some((value) => value.trim())) rows.push(row);
  return rows;
};

const rowsToObjects = (rows) => {
  if (!rows.length) return [];
  const headers = rows[0];
  return rows.slice(1).map((row) =>
    Object.fromEntries(headers.map((header, index) => [header, row[index] || ""]))
  );
};

const buildSkillList = (text) => {
  return String(text || "")
    .split("/")
    .map((skill) => skill.trim())
    .filter(Boolean);
};

const extractMonsterSlug = (imageUrl) => {
  const base = String(imageUrl || "").split("?")[0].split("/").pop() || "";
  return base
    .replace(/\.[^.]+$/, "")
    .replace(/^\d+_/, "")
    .replace(/_\d+g\d+$/i, "");
};

const addCount = (map, name, count) => {
  if (!name || count <= 0) return;
  map.set(name, (map.get(name) || 0) + count);
};

const aggregatePhase = (phase, rows) => {
  const totals = { monster: new Map(), action: new Map(), passive: new Map() };
  let totalAppearances = 0;
  rows.forEach((row) => {
    const count = Number(row["出場回数"] || 0);
    if (count <= 0) return;
    totalAppearances += count;
    addCount(totals.monster, extractMonsterSlug(row.image), count);
    new Set(buildSkillList(row["A(アクティブ)"])).forEach((skill) => {
      addCount(totals.action, skill, count);
    });
    new Set(buildSkillList(row["P(コンパニオン)"])).forEach((skill) => {
      addCount(totals.passive, skill, count);
    });
  });
  return { phase, totalAppearances, totals };
};

const loadPhase = async (phase) => {
  const results = await Promise.all(
    GROUPS.map(async (group) => {
      try {
        const response = await fetch(`./data/${phase}/${group}/skill_list.csv`);
        if (!response.ok) return [];
        return rowsToObjects(parseCsv(await response.text()));
      } catch (error) {
        console.warn(`${phase}期 ${group} の読み込みに失敗しました。`, error);
        return [];
      }
    })
  );
  return aggregatePhase(phase, results.flat());
};

const normalizeText = (value) => {
  return String(value || "").normalize("NFKC").toLowerCase();
};

const getCatalog = () => {
  const names = new Set();
  phaseData.forEach((data) => data.totals[currentType].forEach((count, name) => names.add(name)));
  return Array.from(names).sort((a, b) => a.localeCompare(b, "ja"));
};

const updateOptions = () => {
  itemOptions.innerHTML = "";
  getCatalog().forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    itemOptions.appendChild(option);
  });
};

const getLatestTopNames = () => {
  const latest = phaseData[phaseData.length - 1];
  if (!latest) return [];
  return Array.from(latest.totals[currentType], ([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name, "ja"))
    .slice(0, 5)
    .map((entry) => entry.name);
};

const createSvgElement = (name, attributes = {}) => {
  const element = document.createElementNS(SVG_NS, name);
  Object.entries(attributes).forEach(([key, value]) => element.setAttribute(key, value));
  return element;
};

const getRate = (data, name) => {
  if (!data.totalAppearances) return 0;
  return ((data.totals[currentType].get(name) || 0) / data.totalAppearances) * 100;
};

const renderChart = () => {
  trendChart.innerHTML = "";
  chartWrap.hidden = !selectedNames.length;
  chartNote.hidden = !selectedNames.length;
  if (!selectedNames.length || !phaseData.length) return;

  const width = 1000;
  const height = 500;
  const margin = { top: 30, right: 24, bottom: 58, left: 68 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const allRates = selectedNames.flatMap((name) => phaseData.map((data) => getRate(data, name)));
  const highestRate = Math.max(...allRates, 0);
  const yMax = Math.min(100, Math.max(5, Math.ceil(highestRate / 5) * 5));

  for (let step = 0; step <= 5; step += 1) {
    const rate = (yMax / 5) * step;
    const y = margin.top + plotHeight - (rate / yMax) * plotHeight;
    trendChart.appendChild(
      createSvgElement("line", {
        class: "chart-grid",
        x1: margin.left,
        x2: width - margin.right,
        y1: y,
        y2: y,
      })
    );
    const label = createSvgElement("text", {
      class: "chart-axis-label",
      x: margin.left - 10,
      y: y + 4,
      "text-anchor": "end",
    });
    label.textContent = `${rate.toFixed(rate % 1 ? 1 : 0)}%`;
    trendChart.appendChild(label);
  }

  const labelInterval = phaseData.length > 20 ? 2 : 1;
  phaseData.forEach((data, index) => {
    if (index % labelInterval && index !== phaseData.length - 1) return;
    const x =
      margin.left + (index / Math.max(phaseData.length - 1, 1)) * plotWidth;
    const label = createSvgElement("text", {
      class: "chart-axis-label",
      x,
      y: height - 24,
      "text-anchor": "middle",
    });
    label.textContent = `${data.phase}期`;
    trendChart.appendChild(label);
  });

  selectedNames.forEach((name, seriesIndex) => {
    const color = SERIES_COLORS[seriesIndex];
    const points = phaseData.map((data, index) => {
      const rate = getRate(data, name);
      return {
        data,
        rate,
        x: margin.left + (index / Math.max(phaseData.length - 1, 1)) * plotWidth,
        y: margin.top + plotHeight - (rate / yMax) * plotHeight,
      };
    });
    const path = createSvgElement("path", {
      class: "chart-line",
      stroke: color,
      d: points.map((point, index) => `${index ? "L" : "M"}${point.x},${point.y}`).join(" "),
    });
    trendChart.appendChild(path);
    points.forEach((point) => {
      const circle = createSvgElement("circle", {
        class: "chart-point",
        fill: color,
        cx: point.x,
        cy: point.y,
        r: 5,
      });
      const title = createSvgElement("title");
      const count = point.data.totals[currentType].get(name) || 0;
      title.textContent = `${point.data.phase}期 ${name}: ${point.rate.toFixed(1)}%（${count}/${point.data.totalAppearances}）`;
      circle.appendChild(title);
      trendChart.appendChild(circle);
    });
  });
};

const renderSelectedItems = () => {
  selectedItems.innerHTML = "";
  selectedNames.forEach((name, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "selected-chip";
    button.title = `${name}をグラフから外す`;
    const color = document.createElement("span");
    color.className = "selected-color";
    color.style.background = SERIES_COLORS[index];
    const label = document.createElement("span");
    label.textContent = name;
    const close = document.createElement("span");
    close.textContent = "×";
    button.append(color, label, close);
    button.addEventListener("click", () => {
      selectedNames = selectedNames.filter((selectedName) => selectedName !== name);
      renderSelectedItems();
      renderChart();
    });
    selectedItems.appendChild(button);
  });
};

const resetSelection = () => {
  selectedNames = getLatestTopNames();
  updateOptions();
  renderSelectedItems();
  renderChart();
};

const setType = (type) => {
  if (!TYPES[type] || type === currentType) return;
  currentType = type;
  document.querySelectorAll(".type-tab").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.type === currentType);
  });
  itemInput.value = "";
  resetSelection();
};

typeTabs.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-type]");
  if (button) setType(button.dataset.type);
});

itemPicker.addEventListener("submit", (event) => {
  event.preventDefault();
  const query = normalizeText(itemInput.value.trim());
  if (!query) return;
  const catalog = getCatalog();
  const match =
    catalog.find((name) => normalizeText(name) === query) ||
    catalog.find((name) => normalizeText(name).includes(query));
  if (!match || selectedNames.includes(match)) return;
  if (selectedNames.length >= 5) selectedNames.shift();
  selectedNames.push(match);
  itemInput.value = "";
  renderSelectedItems();
  renderChart();
});

const init = async () => {
  try {
    const response = await fetch("./phase_map.json", { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const json = await response.json();
    const phases = (Array.isArray(json) ? json : json.phases || [])
      .map(Number)
      .filter(Number.isFinite)
      .sort((a, b) => a - b);
    phaseData = await Promise.all(phases.map(loadPhase));
    phaseData = phaseData.filter((data) => data.totalAppearances > 0);
    loadingStatus.textContent = `${phaseData.length}期分を読み込みました。`;
    resetSelection();
  } catch (error) {
    console.error("採用率推移データの読み込みに失敗しました。", error);
    loadingStatus.textContent = "データを読み込めませんでした。";
  }
};

init();
