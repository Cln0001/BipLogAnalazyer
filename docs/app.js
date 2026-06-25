const DATA_URL = "./data.json";
const SUMMARY_KEY = "Summary";

const CLASS_COLORS = {
  Warrior: "C79C6E",
  Paladin: "F58CBA",
  Hunter: "ABD473",
  Rogue: "FFF569",
  Priest: "E8E8E8",
  Shaman: "0070DE",
  Mage: "69CCF0",
  Warlock: "9482C9",
  Druid: "FF7D0A",
};

let DATA = null;
let CURRENT_PHASE = null;

function currentPhaseData() {
  return DATA.phases[CURRENT_PHASE];
}

const WCL_REPORT_BASE = "https://fresh.warcraftlogs.com/reports/";

// Consumables that are technically used by a role sometimes but aren't
// worth showing in that role's table at all (per guild request) — unlike
// excel_export.py's ROLE_RELEVANCE, this hides the row outright rather
// than just blanking zero cells.
// Cosmetic only — mirrors data/phases.json's zone keywords as a short
// display label. Add an entry here when a new phase entry is added there.
const PHASE_LABELS = {
  "2": "SSC/TK",
};

const ROLE_EXCLUDE = {
  "Healer": new Set(["Living/Free Action Potion", "Scroll: Strength Uptime %", "Scroll: Agility Uptime %"]),
  "Caster DPS": new Set(["Living/Free Action Potion", "Scroll: Strength Uptime %", "Scroll: Agility Uptime %"]),
  "Physical DPS": new Set(["Ironshield Potion"]),
};

function fontColorFor(hex) {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.6 ? "#000000" : "#ffffff";
}

function playersForSelection(key) {
  const phase = currentPhaseData();
  if (key === SUMMARY_KEY) return phase.summary.players;
  return phase.logs[key].players;
}

function groupByRole(players) {
  const byRole = new Map();
  for (const p of players) {
    if (!byRole.has(p.role)) byRole.set(p.role, []);
    byRole.get(p.role).push(p);
  }
  const groups = [];
  for (const role of DATA.role_order) {
    const members = byRole.get(role) || [];
    if (members.length === 0) continue;
    members.sort((a, b) => (a.class_name + a.name).localeCompare(b.class_name + b.name));
    groups.push([role, members]);
  }
  return groups;
}

function consumablesOf(player, isSummary) {
  return isSummary ? player.consumables_avg : player.consumables;
}

function groupNamesByCategory(names) {
  const byCategory = new Map();
  for (const name of names) {
    const category = DATA.consumable_categories[name] || "Other";
    if (!byCategory.has(category)) byCategory.set(category, []);
    byCategory.get(category).push(name);
  }
  const groups = [];
  for (const category of DATA.category_order) {
    const inCategory = byCategory.get(category);
    if (!inCategory || inCategory.length === 0) continue;
    inCategory.sort();
    groups.push([category, inCategory]);
  }
  return groups;
}

function buildRoleTable(role, members, isSummary) {
  const excluded = ROLE_EXCLUDE[role];
  const names = new Set();
  for (const p of members) {
    for (const n of Object.keys(consumablesOf(p, isSummary))) {
      if (!excluded || !excluded.has(n)) names.add(n);
    }
  }
  const categoryGroups = groupNamesByCategory(names);
  const colCount = members.length + 1;

  const table = document.createElement("table");
  table.className = "role-table";

  const caption = document.createElement("caption");
  caption.textContent = role;
  table.appendChild(caption);

  const headRow = document.createElement("tr");
  const cornerCell = document.createElement("th");
  cornerCell.textContent = "Name";
  cornerCell.className = "name-col";
  headRow.appendChild(cornerCell);
  for (const p of members) {
    const th = document.createElement("th");
    th.textContent = p.name;
    th.className = "player-col";
    const color = CLASS_COLORS[p.class_name];
    if (color) {
      th.style.backgroundColor = "#" + color;
      th.style.color = fontColorFor(color);
    }
    headRow.appendChild(th);
  }
  table.appendChild(headRow);

  if (isSummary) {
    const attendedRow = document.createElement("tr");
    attendedRow.appendChild(rowLabelCell("Logs Attended"));
    for (const p of members) attendedRow.appendChild(plainCell(p.logs_attended));
    table.appendChild(attendedRow);
  }

  for (const [category, categoryNames] of categoryGroups) {
    const sectionRow = document.createElement("tr");
    sectionRow.className = "category-row";
    const sectionCell = document.createElement("th");
    sectionCell.colSpan = colCount;
    sectionCell.textContent = isSummary ? `⌀ ${category} per Raid` : category;
    sectionRow.appendChild(sectionCell);
    table.appendChild(sectionRow);

    for (const name of categoryNames) {
      const row = document.createElement("tr");
      row.appendChild(rowLabelCell(name));
      for (const p of members) {
        const value = consumablesOf(p, isSummary)[name] ?? 0;
        row.appendChild(plainCell(value === 0 ? "" : value));
      }
      table.appendChild(row);
    }

    // Summing percentages (every "... Uptime %" row) produces a meaningless
    // number — e.g. Flask 19% + Elixir 20% + Well Fed 38% isn't a 77%
    // anything. Skip the Total row entirely when every row in this category
    // is a percentage metric; only render it for real summable counts.
    if (!categoryNames.every((name) => name.endsWith("Uptime %"))) {
      const totalRow = document.createElement("tr");
      totalRow.className = "total-row";
      totalRow.appendChild(rowLabelCell("Total"));
      for (const p of members) {
        const consumables = consumablesOf(p, isSummary);
        const total = categoryNames.reduce((sum, name) => sum + (consumables[name] ?? 0), 0);
        const cell = plainCell(total === 0 ? "" : Math.round(total * 10) / 10);
        totalRow.appendChild(cell);
      }
      table.appendChild(totalRow);
    }
  }

  return table;
}

function rowLabelCell(text) {
  const th = document.createElement("th");
  th.scope = "row";
  th.className = "name-col";
  th.textContent = text;
  return th;
}

function plainCell(value) {
  const td = document.createElement("td");
  td.className = "player-col";
  td.textContent = value;
  return td;
}

function render(key) {
  const container = document.getElementById("table-container");
  container.innerHTML = "";

  const isSummary = key === SUMMARY_KEY;
  const players = playersForSelection(key);
  const groups = groupByRole(players);

  const phase = currentPhaseData();

  const meta = document.getElementById("meta");
  meta.textContent = isSummary ? `${phase.log_order.length} logs` : "";

  const link = document.getElementById("report-link");
  if (isSummary) {
    link.style.display = "none";
  } else {
    const entry = phase.logs[key];
    link.href = WCL_REPORT_BASE + entry.report_code;
    link.textContent = "View on Warcraft Logs ↗";
    link.style.display = "inline";
  }

  for (const [role, members] of groups) {
    container.appendChild(buildRoleTable(role, members, isSummary));
  }

  fitColumnsToContent(container, ".name-col");
  fitColumnsToContent(container, ".player-col");
}

function fitColumnsToContent(container, selector) {
  const cells = container.querySelectorAll(selector);
  let maxWidth = 0;
  for (const cell of cells) maxWidth = Math.max(maxWidth, cell.scrollWidth);
  for (const cell of cells) cell.style.width = `${maxWidth}px`;
}

function populateLogPicker() {
  const picker = document.getElementById("log-picker");
  picker.innerHTML = "";

  const summaryOption = document.createElement("option");
  summaryOption.value = SUMMARY_KEY;
  summaryOption.textContent = SUMMARY_KEY;
  picker.appendChild(summaryOption);

  const phase = currentPhaseData();
  for (const key of [...phase.log_order].reverse()) {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = phase.logs[key].log_date;
    picker.appendChild(option);
  }
}

function populatePhasePicker() {
  const picker = document.getElementById("phase-picker");
  picker.innerHTML = "";

  for (const phase of DATA.phase_order) {
    const option = document.createElement("option");
    option.value = phase;
    const label = PHASE_LABELS[phase];
    option.textContent = label ? `Phase ${phase} (${label})` : `Phase ${phase}`;
    picker.appendChild(option);
  }

  // Default to the highest/most recent phase (last entry in phase_order,
  // which is sorted numerically-then-alphabetically by the backend).
  CURRENT_PHASE = DATA.phase_order[DATA.phase_order.length - 1];
  picker.value = CURRENT_PHASE;

  picker.addEventListener("change", (e) => {
    CURRENT_PHASE = e.target.value;
    populateLogPicker();
    render(SUMMARY_KEY);
  });
}

async function init() {
  const res = await fetch(DATA_URL, { cache: "no-store" });
  DATA = await res.json();
  populatePhasePicker();
  populateLogPicker();
  document.getElementById("log-picker").addEventListener("change", (e) => render(e.target.value));
  render(SUMMARY_KEY);
}

init();
