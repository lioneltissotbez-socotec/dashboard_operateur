/* =========================================================================
   VARIABLES GLOBALES
   ========================================================================= */

let ALL_OPERATORS = [];
let LOCATION_MAP = {};
let HAS_LOCATION = false;

let VIEW_MODE = "list";

let chartBubble = null;
let chartPie = null;

const fileInput = document.getElementById("excelFile");
const locationInput = document.getElementById("excelLocation");

const tbody = document.querySelector("#dataTable tbody");

const filterExpiry = document.getElementById("filterExpiry");
const filterCert = document.getElementById("filterCert");
const filterPole = document.getElementById("filterPole");
const filterSection = document.getElementById("filterSection");

const searchInput = document.getElementById("searchOperator");
const searchSelect = document.getElementById("searchOperatorSelect");

const domainCheckboxes = document.querySelectorAll('.domain-toggle input[type="checkbox"]');

const pieContainer = document.getElementById("pieContainer");
const bubbleContainer = document.getElementById("bubbleContainer");
const tableContainer = document.getElementById("tableContainer");
const cardsContainer = document.getElementById("cardsContainer");

const kpiTotal = document.getElementById("kpiTotal");
const kpiFiltered = document.getElementById("kpiFilteredTotal");
const kpiExpired = document.getElementById("kpiFilteredExpired");
const kpiFull = document.getElementById("kpiFilteredFull");

/* =========================================================================
   DOMAINS CONFIG
   ========================================================================= */

const DOMAINS = [
  {label:"Amiante",     key:"amiante", isCore:true,  cert:10, num:11, deb:12, fin:13},
  {label:"CREP",        key:"crep",    isCore:true,  cert:15, num:16, deb:17, fin:18},
  {label:"Termites",    key:"term",    isCore:true,  cert:20, num:21, deb:22, fin:23},
  {label:"DPE Mention", key:"dpem",    isCore:true,  cert:25, num:26, deb:27, fin:28},
  {label:"Gaz",         key:"gaz",     isCore:true,  cert:30, num:31, deb:32, fin:33},
  {label:"Élec",        key:"elec",    isCore:true,  cert:35, num:36, deb:37, fin:38},
  {label:"DPE Indiv",   key:"dpei",    isCore:false, cert:42, num:43, deb:44, fin:45},
  {label:"Audit",       key:"audit",   isCore:false, audit:48}
];

/* =========================================================================
   UTILITAIRES
   ========================================================================= */

function parseExcelDate(v) {
  if (!v) return null;

  if (v instanceof Date) return v;

  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    return new Date(d.y, d.m - 1, d.d);
  }

  if (typeof v === "string") {
    const p = v.split("/");
    if (p.length === 3) {
      return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
  }
  return null;
}

function isValidDate(dt) {
  if (!dt) return false;
  const today = new Date();
  today.setHours(0,0,0,0);
  const d = new Date(dt.getTime());
  d.setHours(0,0,0,0);
  return d >= today;
}

function normalizeString(s) {
  return (s || "")
    .toString()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9 ]/g," ")
    .replace(/\s+/g," ")
    .trim()
    .toUpperCase();
}

function getGeoKeyFromRow(row) {
  return normalizeString((row[1] || "") + " " + (row[2] || ""));
}

function getGeoKeyFromOperator(op) {
  const nom = op.row[2] || "";
  const prenom = op.row[3] || "";
  const direct = normalizeString(nom + " " + prenom);
  const reverse = normalizeString(prenom + " " + nom);
  if (LOCATION_MAP[direct]) return direct;
  if (LOCATION_MAP[reverse]) return reverse;
  return direct;
}

/* =========================================================================
   CHARGEMENT FICHIERS
   ========================================================================= */

document.getElementById("openFileBtn").addEventListener("click", () => fileInput.click());
document.getElementById("openLocationBtn").addEventListener("click", () => locationInput.click());

fileInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => loadOperators(evt.target.result);
  reader.readAsBinaryString(file);
});

locationInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = evt => loadLocation(evt.target.result);
  reader.readAsBinaryString(file);
});

/* =========================================================================
   PARSING LOCATION
   ========================================================================= */

function loadLocation(data) {
  const wb = XLSX.read(data, {type:"binary"});
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1});

  LOCATION_MAP = {};

  rows.slice(1).forEach(row => {
    const key = getGeoKeyFromRow(row);
    LOCATION_MAP[key] = {
      email: row[3] || "",
      section: row[4] || "—",
      pole: row[5] || "—",
      manager: ((row[6] || "") + " " + (row[7] || "")).trim()
    };
  });

  HAS_LOCATION = true;

  attachLocationToOperators();
  applyFilters();
}

/* =========================================================================
   PARSING OPERATORS
   ========================================================================= */

function loadOperators(data) {
  ALL_OPERATORS = [];

  const wb = XLSX.read(data, {type:"binary", cellDates:true});
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {header:1});

  rows.slice(1).forEach(row => {
    const name = (row[1] || "").toString().trim();
    if (!name) return;

    const op = {name, row, domains:{}, pole:"—", section:"—", manager:"—", email:""};

    DOMAINS.forEach(domain => {
      let status = "none";
      let org="—", num="—", deb=null, fin=null;

      if (domain.audit != null) {
        if (row[domain.audit]) {
          status = "valid";
          org = "OK";
        }
      } else {
        org = row[domain.cert] || "—";
        num = row[domain.num] || "—";
        deb = parseExcelDate(row[domain.deb]);
        fin = parseExcelDate(row[domain.fin]);
        if (!org && !num && !deb && !fin) status = "none";
        else if (fin && !isValidDate(fin)) status = "expired";
        else status = "valid";
      }

      op.domains[domain.key] = {status, org, num, deb, fin};
    });

    ALL_OPERATORS.push(op);
  });

  if (HAS_LOCATION) attachLocationToOperators();

  populateFilters();
  applyFilters();
}

function attachLocationToOperators() {
  ALL_OPERATORS.forEach(op => {
    const key = getGeoKeyFromOperator(op);
    const loc = LOCATION_MAP[key];
    if (!loc) return;
    op.email = loc.email;
    op.section = loc.section;
    op.pole = loc.pole;
    op.manager = loc.manager;
  });
}

/* =========================================================================
   FILTRES
   ========================================================================= */

function populateFilters() {
  const certs = new Set();
  const poles = new Set();
  const sections = new Set();

  ALL_OPERATORS.forEach(op => {
    DOMAINS.forEach(d => {
      const org = op.domains[d.key].org;
      if (org && org !== "—") certs.add(org);
    });
    if (op.pole !== "—") poles.add(op.pole);
    if (op.section !== "—") sections.add(op.section);
  });

  filterCert.innerHTML = "<option value='ALL'>Tous</option>";
  Array.from(certs).sort().forEach(c => {
    filterCert.innerHTML += `<option value="${c}">${c}</option>`;
  });

  filterPole.innerHTML = "<option value='ALL'>Tous</option>";
  Array.from(poles).sort().forEach(p => {
    filterPole.innerHTML += `<option value="${p}">${p}</option>`;
  });

  filterSection.innerHTML = "<option value='ALL'>Toutes</option>";
  Array.from(sections).sort().forEach(s => {
    filterSection.innerHTML += `<option value="${s}">${s}</option>`;
  });

  searchSelect.innerHTML = "<option value='ALL'>Tous</option>";
  ALL_OPERATORS
    .map(o => o.name)
    .sort()
    .forEach(n => {
      searchSelect.innerHTML += `<option value="${n}">${n}</option>`;
    });
}

/* =========================================================================
   APPLICATION DES FILTRES GLOBAUX
   ========================================================================= */

function applyFilters() {
  let ops = ALL_OPERATORS.slice();

  const txt = searchInput.value.trim().toLowerCase();
  if (txt) ops = ops.filter(o => o.name.toLowerCase().includes(txt));

  if (searchSelect.value !== "ALL") {
    ops = ops.filter(o => o.name === searchSelect.value);
  }

  if (filterCert.value !== "ALL") {
    ops = ops.filter(o =>
      DOMAINS.some(d => o.domains[d.key].org === filterCert.value)
    );
  }

  if (filterPole.value !== "ALL") ops = ops.filter(o => o.pole === filterPole.value);
  if (filterSection.value !== "ALL") ops = ops.filter(o => o.section === filterSection.value);

  if (parseInt(filterExpiry.value) > 0) {
    const months = parseInt(filterExpiry.value);
    const today = new Date();
    const limit = new Date(today.getFullYear(), today.getMonth()+months, today.getDate());
    ops = ops.filter(o => {
      return DOMAINS.some(d => {
        const fin = o.domains[d.key].fin;
        return fin && fin >= today && fin <= limit;
      });
    });
  }

  updateKPI(ops);
  renderTable(ops);
  renderCards(ops);
  renderBubbleChart(ops);
  renderPieChart(ops);

  updateVisibility();
}

function updateKPI(ops) {
  kpiTotal.textContent = ALL_OPERATORS.length;
  kpiFiltered.textContent = ops.length;
  kpiExpired.textContent = ops.filter(o =>
    DOMAINS.some(d => o.domains[d.key].status === "expired")
  ).length;
  kpiFull.textContent = ops.filter(o =>
    DOMAINS.filter(d => d.isCore).every(d => o.domains[d.key].status === "valid")
  ).length;
}

/* =========================================================================
   TABLEAU
   ========================================================================= */

function renderTable(ops) {
  tbody.innerHTML = "";
  ops.forEach(op => {
    const tr = document.createElement("tr");
    let html = `<td>${op.name}</td>`;
    DOMAINS.forEach(d => {
      const info = op.domains[d.key];
      const cls = info.status;
      const org = info.org || "—";
      html += `
        <td>
          <div class="cell-content">
            <span class="dot ${cls}"></span>
            <span class="cert-org ${cls}">${org}</span>
          </div>
        </td>`;
    });
    html += `<td>${op.pole}</td><td>${op.section}</td><td>${op.manager}</td>`;
    tr.innerHTML = html;
    tbody.appendChild(tr);
  });
}

/* =========================================================================
   CARTES (MODE CARDS)
   ========================================================================= */

function renderCards(ops) {
  cardsContainer.innerHTML = "";

  if (VIEW_MODE !== "cards") return;

  DOMAINS.forEach(domain => {
    const list = ops.filter(o => o.domains[domain.key].status === "valid");
    if (!list.length) return;

    const div = document.createElement("div");
    div.className = "memoire-card";
    div.innerHTML = `
      <div class="memoire-card-header">
        <span class="memoire-card-title">${domain.label}</span>
        <span class="memoire-card-badge">${list.length} certifiés</span>
      </div>
      <div class="memoire-names">
        ${list.map(o => `<span class="memoire-name-chip">${o.name}</span>`).join("")}
      </div>`;
    cardsContainer.appendChild(div);
  });
}

/* =========================================================================
   BUBBLE CHART
   ========================================================================= */

function renderBubbleChart(ops) {
  if (VIEW_MODE !== "bubbles") {
    bubbleContainer.style.display = "none";
    if (chartBubble) chartBubble.destroy();
    return;
  }

  bubbleContainer.style.display = "block";

  const counts = DOMAINS.map(d => {
    return ops.filter(o => o.domains[d.key].status === "valid").length;
  });

  const max = Math.max(...counts, 1);

  const dataset = DOMAINS.map((d, i) => ({
    label: `${d.label} (${counts[i]})`,
    data: [{x:(i+1)*10, y:50, r:(counts[i] / max)*30 + 10}],
    backgroundColor: CHART_COLORS[i % CHART_COLORS.length]
  }));

  if (chartBubble) chartBubble.destroy();

  chartBubble = new Chart(document.getElementById("bubbleChart").getContext("2d"), {
    type:"bubble",
    data:{datasets:dataset},
    options:{
      responsive:true,
      plugins:{legend:{position:"bottom"}},
      scales:{x:{display:false,min:0,max:100},y:{display:false,min:0,max:100}}
    }
  });
}

/* =========================================================================
   PIE CHART (Camembert global)
   ========================================================================= */

const CHART_COLORS = [
  "#2ecc71","#3498db","#9b59b6","#e67e22",
  "#e74c3c","#16a085","#f1c40f","#34495e"
];

function renderPieChart(ops) {
  if (VIEW_MODE !== "pie") {
    pieContainer.style.display = "none";
    if (chartPie) chartPie.destroy();
    return;
  }

  pieContainer.style.display = "block";

  // Uniquement VALID + near6 + near12
  const counts = DOMAINS.map(d => {
    return ops.filter(o => {
      const st = o.domains[d.key].status;
      return st === "valid" || st === "near6" || st === "near12";
    }).length;
  });

  if (chartPie) chartPie.destroy();

  chartPie = new Chart(document.getElementById("pieChart").getContext("2d"), {
    type: "pie",
    data: {
      labels: DOMAINS.map((d,i)=> `${d.label} (${counts[i]})`),
      datasets: [{
        data: counts,
        backgroundColor: CHART_COLORS
      }]
    },
    options: {
      responsive:true,
      plugins:{legend:{position:"right"}}
    }
  });
}

/* =========================================================================
   GESTION DES MODES
   ========================================================================= */

document.getElementById("layoutListBtn").addEventListener("click", () => {
  VIEW_MODE = "list";
  updateVisibility();
  applyFilters();
});

document.getElementById("layoutCardsBtn").addEventListener("click", () => {
  VIEW_MODE = "cards";
  updateVisibility();
  applyFilters();
});

document.getElementById("layoutBubblesBtn").addEventListener("click", () => {
  VIEW_MODE = "bubbles";
  updateVisibility();
  applyFilters();
});

document.getElementById("layoutPieBtn").addEventListener("click", () => {
  VIEW_MODE = "pie";
  updateVisibility();
  applyFilters();
});

function updateVisibility() {
  tableContainer.style.display  = (VIEW_MODE === "list")   ? "block" : "none";
  cardsContainer.style.display  = (VIEW_MODE === "cards")  ? "grid"  : "none";
  bubbleContainer.style.display = (VIEW_MODE === "bubbles")? "block" : "none";
  pieContainer.style.display    = (VIEW_MODE === "pie")    ? "block" : "none";
}
