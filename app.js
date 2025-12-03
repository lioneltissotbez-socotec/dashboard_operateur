/* ==========
   Utilitaires dates
   ========== */

function parseExcelDate(v) {
  if (!v) return null;

  if (v instanceof Date) return v;

  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return new Date(d.y, d.m - 1, d.d);
  }

  if (typeof v === "string") {
    const parts = v.split("/");
    if (parts.length === 3) {
      const dd = parseInt(parts[0], 10);
      const mm = parseInt(parts[1], 10);
      const yyyy = parseInt(parts[2], 10);
      if (!isNaN(dd) && !isNaN(mm) && !isNaN(yyyy)) {
        return new Date(yyyy, mm - 1, dd);
      }
    }
  }
  return null;
}

function isValidDate(dateFin) {
  const d = parseExcelDate(dateFin);
  if (!d) return false;
  const today = new Date();
  today.setHours(0,0,0,0);
  d.setHours(0,0,0,0);
  return d >= today;
}

/* ==========
   Normalisation / clés de jointure
   ========== */

function normalizeString(str) {
  if (!str) return "";

  return str
    .toString()
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[-'_]+/g, " ")
    .replace(/[^A-Za-z0-9 ]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

// Clé depuis une ligne du fichier géographique : Nom (B) + Prénom (C)
function getGeoKeyFromRow(row) {
  const nom = row[1] || "";
  const prenom = row[2] || "";
  return normalizeString(nom + " " + prenom);
}

// Clé depuis un opérateur (en priorité colonnes C = NOM, D = Prénom)
function getGeoKeyFromOperator(op) {
  const r = op.row || [];
  const nom = r[2] || "";
  const prenom = r[3] || "";

  const direct = normalizeString(nom + " " + prenom);     // NOM PRÉNOM
  const reversed = normalizeString(prenom + " " + nom);   // PRÉNOM NOM

  if (nom || prenom) {
    if (LOCATION_MAP[direct]) return direct;
    if (LOCATION_MAP[reversed]) return reversed;
    return direct;
  }

  const parts = (op.name || "").split(" ").filter(x => x.trim() !== "");
  if (parts.length >= 2) {
    const last = parts[parts.length - 1];
    const first = parts.slice(0, -1).join(" ");
    const p_direct = normalizeString(last + " " + first);
    const p_reversed = normalizeString(first + " " + last);

    if (LOCATION_MAP[p_direct]) return p_direct;
    if (LOCATION_MAP[p_reversed]) return p_reversed;

    return p_direct;
  }

  return normalizeString(op.name || "");
}

/* ==========
   Config colonnes / domaines
   ========== */

const COL = {
  NAME:1,

  AMI_CERT:10, AMI_NUM:11, AMI_DEB:12, AMI_FIN:13,
  CREP_CERT:15, CREP_NUM:16, CREP_DEB:17, CREP_FIN:18,
  TERM_CERT:20, TERM_NUM:21, TERM_DEB:22, TERM_FIN:23,
  DPEM_CERT:25, DPEM_NUM:26, DPEM_DEB:27, DPEM_FIN:28,
  GAZ_CERT:30, GAZ_NUM:31, GAZ_DEB:32, GAZ_FIN:33,
  ELEC_CERT:35, ELEC_NUM:36, ELEC_DEB:37, ELEC_FIN:38,
  DPEI_CERT:42, DPEI_NUM:43, DPEI_DEB:44, DPEI_FIN:45,
  AUDIT:48
};

const DOMAINS = [
  {label:"Amiante",     key:"amiante", isCore:true,
    cert:COL.AMI_CERT,  num:COL.AMI_NUM,  deb:COL.AMI_DEB,  fin:COL.AMI_FIN},
  {label:"CREP",        key:"crep",    isCore:true,
    cert:COL.CREP_CERT, num:COL.CREP_NUM, deb:COL.CREP_DEB, fin:COL.CREP_FIN},
  {label:"Termites",    key:"term",    isCore:true,
    cert:COL.TERM_CERT, num:COL.TERM_NUM, deb:COL.TERM_DEB, fin:COL.TERM_FIN},
  {label:"DPE Mention", key:"dpem",    isCore:true,
    cert:COL.DPEM_CERT, num:COL.DPEM_NUM, deb:COL.DPEM_DEB, fin:COL.DPEM_FIN},
  {label:"Gaz",         key:"gaz",     isCore:true,
    cert:COL.GAZ_CERT,  num:COL.GAZ_NUM,  deb:COL.GAZ_DEB,  fin:COL.GAZ_FIN},
  {label:"Élec",        key:"elec",    isCore:true,
    cert:COL.ELEC_CERT, num:COL.ELEC_NUM, deb:COL.ELEC_DEB, fin:COL.ELEC_FIN},
  {label:"DPE Indiv",   key:"dpei",    isCore:false,
    cert:COL.DPEI_CERT, num:COL.DPEI_NUM, deb:COL.DPEI_DEB, fin:COL.DPEI_FIN},
  {label:"Audit",       key:"audit",   isCore:false,
    audit:COL.AUDIT}
];

/* ==========
   Références DOM et variables globales
   ========== */

const fileInput = document.getElementById("excelFile");
const locationFileInput = document.getElementById("excelLocation");

const filterSelect = document.getElementById("filterExpiry");
const filterCertSelect = document.getElementById("filterCert");
const filterPoleSelect = document.getElementById("filterPole");
const filterSectionSelect = document.getElementById("filterSection");

const graphModeSelect = document.getElementById("graphMode");
const loadingBarWrapper = document.getElementById("loadingBarWrapper");
const loadingBar = document.getElementById("loadingBar");
const tbody = document.querySelector("#dataTable tbody");
const kpiTotal = document.getElementById("kpiTotal");
const kpiExpired = document.getElementById("kpiExpired");
const kpiFull = document.getElementById("kpiFull");
const chartCanvas = document.getElementById("chart");
const exportVisualBtn = document.getElementById("exportVisualBtn");
const exportModal = document.getElementById("exportModal");
const exportTableContainer = document.getElementById("exportTableContainer");
const copyExportBtn = document.getElementById("copyExportBtn");
const closeExportBtn = document.getElementById("closeExportBtn");
const searchOperatorInput = document.getElementById("searchOperator");
const searchOperatorSelect = document.getElementById("searchOperatorSelect");
const toggleGeoColsCheckbox = document.getElementById("toggleGeoCols");
const exportEmailsBtn = document.getElementById("exportEmailsBtn");
const toggleCert = document.getElementById("toggleCert");


const domainCheckboxes = document.querySelectorAll('.domain-toggle input[type="checkbox"]');

function getSelectLabel(selectEl, fallback){
  if (!selectEl) return fallback;
  const opt = selectEl.options[selectEl.selectedIndex];
  return opt ? opt.textContent.trim() : fallback;
}

function setChartEmptyState(container, canvas, emptyId, isEmpty, message) {
  if (!container || !canvas) return;

  let emptyEl = document.getElementById(emptyId);
  if (!emptyEl) {
    emptyEl = document.createElement("div");
    emptyEl.id = emptyId;
    emptyEl.className = "chart-empty";
    container.insertBefore(emptyEl, canvas);
  }

  emptyEl.textContent = message || "Aucune donnée à afficher";
  emptyEl.style.display = isEmpty ? "block" : "none";
  canvas.style.display = isEmpty ? "none" : "block";
}

const chartValueLabelPlugin = {
  id: "chartValueLabel",
  afterDatasetDraw(chart, args, pluginOptions){
    const {ctx} = chart;
    ctx.save();
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.font = (pluginOptions && pluginOptions.font) || "12px 'Inter', 'Segoe UI', Arial, sans-serif";

    if (chart.config.type === "pie") {
      const dataset = chart.data.datasets[0];
      const meta = chart.getDatasetMeta(0);
      meta.data.forEach((arc, idx) => {
        const value = dataset.data[idx];
        if (!value) return;
        const {x, y} = arc.tooltipPosition();
        ctx.fillStyle = (pluginOptions && pluginOptions.pieColor) || "#0f172a";
        ctx.fillText(value, x, y);
      });
    }

    if (chart.config.type === "bubble") {
      chart.data.datasets.forEach((dataset, datasetIndex) => {
        const meta = chart.getDatasetMeta(datasetIndex);
        meta.data.forEach((pt, i) => {
          const raw = dataset.data[i] || {};
          const value = raw.count ?? dataset._count ?? raw.r;
          if (!value && value !== 0) return;
          const {x, y} = pt.tooltipPosition();
          ctx.fillStyle = (pluginOptions && pluginOptions.bubbleColor) || "#ffffff";
          ctx.fillText(value, x, y);
        });
      });
    }

    ctx.restore();
  }
};

function updateDomainVisibility() {
  const table = document.getElementById("dataTable");
  if (!table) return;
  const headRow = table.tHead ? table.tHead.rows[0] : null;
  if (!headRow) return;

  const visibleMap = {};
  domainCheckboxes.forEach(cb => {
    visibleMap[cb.value] = cb.checked;
  });

  const ths = headRow.cells;

  DOMAINS.forEach((domain, idx) => {
    const colIndex = 1 + idx; // 0 = opérateur, ensuite domaines
    const visible = visibleMap[domain.key] !== false;
    const displayValue = visible ? "" : "none";

    if (ths[colIndex]) ths[colIndex].style.display = displayValue;

    Array.from(table.tBodies[0].rows).forEach(row => {
      if (row.cells[colIndex]) row.cells[colIndex].style.display = displayValue;
    });
  });
}

domainCheckboxes.forEach(cb => {
  cb.addEventListener("change", () => {
    updateDomainVisibility();
    applyFilters();
  });
});


function renderMemoirePanels(ops) {
  const container = document.getElementById("memoirePanels");
  if (!container) return;

  const isMemoire = (typeof VIEW_MODE !== "undefined" && VIEW_MODE === "memoire");
  if (!isMemoire) {
    container.innerHTML = "";
    return;
  }

  // Carte par domaine, seulement si au moins un opérateur avec certif valide
  const domainVisibleMap = {};
  if (typeof domainCheckboxes !== "undefined") {
    domainCheckboxes.forEach(cb => {
      domainVisibleMap[cb.value] = cb.checked;
    });
  }

  let htmlPanels = "";

  DOMAINS.forEach(domain => {
    // Respecter les domaines masqués
    if (domainVisibleMap[domain.key] === false) return;

    const names = [];

    ops.forEach(op => {
      const d = op.domains[domain.key];
      if (!d) return;
      // En mémoires : on ne retient que les certifs valides (y compris proches)
      if (d.status === "valid") {
        names.push(op.name);
      }
    });

    if (names.length === 0) return;

    const uniqueNames = Array.from(new Set(names)).sort((a,b) => a.localeCompare(b));

    htmlPanels += '<div class="memoire-card">';
    htmlPanels +=   '<div class="memoire-card-header">';
    htmlPanels +=     '<div class="memoire-card-title">' + domain.label + '</div>';
    htmlPanels +=     '<div class="memoire-card-badge">' + uniqueNames.length + ' certifiés</div>';
    htmlPanels +=   '</div>';

    if (uniqueNames.length) {
      htmlPanels += '<div class="memoire-names">';
      uniqueNames.forEach(name => {
        htmlPanels += '<span class="memoire-name-chip">' + name + '</span>';
      });
      htmlPanels += '</div>';
    } else {
      htmlPanels += '<div class="memoire-names-empty">Aucun opérateur certifié</div>';
    }

    htmlPanels += '</div>';
  });

  container.innerHTML = htmlPanels;
}

function updatePieChart(ops) {
  const container = document.getElementById("pieContainer");
  const canvas = document.getElementById("pieChart");
  const contextInfo = document.getElementById("pieContext");
  if (!container || !canvas) return;

  if (contextInfo) {
    const poleLabel = getSelectLabel(filterPoleSelect, "Tous");
    const sectionLabel = getSelectLabel(filterSectionSelect, "Toutes");
    contextInfo.textContent = `Pôle : ${poleLabel} • Section : ${sectionLabel}`;
  }

  const ctx = canvas.getContext("2d");

  const visibleMap = {};
  domainCheckboxes.forEach(cb => {
    visibleMap[cb.value] = cb.checked;
  });

  const labels = [];
  const data = [];

  DOMAINS.forEach(domain => {
    if (visibleMap[domain.key] === false) return;

    let count = 0;
    ops.forEach(op => {
      const d = op.domains[domain.key];
      if (!d) return;
      if (d.status === "valid") count++;
    });

    labels.push(domain.label);
    data.push(count);
  });

  const isEmpty = labels.length === 0;
  setChartEmptyState(container, canvas, "pieEmptyState", isEmpty, "Aucune donnée pour le camembert");
  if (isEmpty) {
    if (pieChartInstance) pieChartInstance.destroy();
    return;
  }

  const palette = [
    "#2ecc71",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#e74c3c",
    "#16a085",
    "#f1c40f",
    "#34495e"
  ];

  if (pieChartInstance) {
    pieChartInstance.destroy();
  }

  pieChartInstance = new Chart(ctx, {
    type: "pie",
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, idx) => palette[idx % palette.length])
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.label || "";
              const value = context.parsed || 0;
              return `${label}: ${value}`;
            }
          }
        }
      }
    }
  });
}

function updateBubbleChart(ops) {
  const container = document.getElementById("bubbleContainer");
  const canvas = document.getElementById("bubbleChart");
  if (!container || !canvas) return;

  const ctx = canvas.getContext("2d");

  const visibleMap = {};
  domainCheckboxes.forEach(cb => {
    visibleMap[cb.value] = cb.checked;
  });

  const labels = [];
  const data = [];

  DOMAINS.forEach(domain => {
    if (visibleMap[domain.key] === false) return;

    let count = 0;
    ops.forEach(op => {
      const d = op.domains[domain.key];
      if (!d) return;
      if (d.status === "valid") count++;
    });

    labels.push(domain.label);
    data.push(count);
  });

  if (!labels.length) {
    if (pieChartInstance) pieChartInstance.destroy();
    return;
  }

  const palette = [
    "#2ecc71",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#e74c3c",
    "#16a085",
    "#f1c40f",
    "#34495e"
  ];

  if (pieChartInstance) {
    pieChartInstance.destroy();
  }

  pieChartInstance = new Chart(ctx, {
    type: "pie",
    plugins: [chartValueLabelPlugin],
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, idx) => palette[idx % palette.length])
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.label || "";
              const value = context.parsed || 0;
              return `${label}: ${value}`;
            }
          }
        },
        chartValueLabel: {
          pieColor: "#0f172a"
        }
      }
    }
  });
}

function updateBubbleChart(ops) {
  const container = document.getElementById("bubbleContainer");
  const canvas = document.getElementById("bubbleChart");
  const contextInfo = document.getElementById("bubbleContext");
  if (!container || !canvas) return;

  if (contextInfo) {
    const poleLabel = getSelectLabel(filterPoleSelect, "Tous");
    const sectionLabel = getSelectLabel(filterSectionSelect, "Toutes");
    contextInfo.textContent = `Pôle : ${poleLabel} • Section : ${sectionLabel}`;
  }

  const ctx = canvas.getContext("2d");

  const visibleMap = {};
  domainCheckboxes.forEach(cb => {
    visibleMap[cb.value] = cb.checked;
  });

  const labels = [];
  const data = [];

  DOMAINS.forEach(domain => {
    if (visibleMap[domain.key] === false) return;

    let count = 0;
    ops.forEach(op => {
      const d = op.domains[domain.key];
      if (!d) return;
      if (d.status === "valid") count++;
    });

    labels.push(domain.label);
    data.push(count);
  });

  if (!labels.length) {
    if (pieChartInstance) pieChartInstance.destroy();
    setChartEmptyState(container, canvas, "pieChartEmpty", true, "Aucun domaine sélectionné");
    return;
  }

  const total = data.reduce((sum, val) => sum + val, 0);
  if (total === 0) {
    if (pieChartInstance) pieChartInstance.destroy();
    const ctx2d = canvas.getContext("2d");
    if (ctx2d) ctx2d.clearRect(0, 0, canvas.width, canvas.height);
    setChartEmptyState(container, canvas, "pieChartEmpty", true, "Aucune donnée valide à afficher");
    return;
  }

  setChartEmptyState(container, canvas, "pieChartEmpty", false);

  const palette = [
    "#2ecc71",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#e74c3c",
    "#16a085",
    "#f1c40f",
    "#34495e"
  ];

  if (pieChartInstance) {
    pieChartInstance.destroy();
  }

  pieChartInstance = new Chart(ctx, {
    type: "pie",
    plugins: [chartValueLabelPlugin],
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, idx) => palette[idx % palette.length])
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom"
        },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.label || "";
              const value = context.parsed || 0;
              return `${label}: ${value}`;
            }
          }
        },
        chartValueLabel: {
          pieColor: "#0f172a"
        }
      }
    }
  });
}

function updateBubbleChart(ops) {
  const container = document.getElementById("bubbleContainer");
  const canvas = document.getElementById("bubbleChart");
  const contextInfo = document.getElementById("bubbleContext");
  if (!container || !canvas) return;

  if (contextInfo) {
    const poleLabel = getSelectLabel(filterPoleSelect, "Tous");
    const sectionLabel = getSelectLabel(filterSectionSelect, "Toutes");
    contextInfo.textContent = `Pôle : ${poleLabel} • Section : ${sectionLabel}`;
  }

  const ctx = canvas.getContext("2d");

  const visibleMap = {};
  domainCheckboxes.forEach(cb => {
    visibleMap[cb.value] = cb.checked;
  });

  const counts = [];
  DOMAINS.forEach((domain, idx) => {
    if (visibleMap[domain.key] === false) return;

    let count = 0;
    ops.forEach(op => {
      const d = op.domains[domain.key];
      if (!d) return;
      if (d.status === "valid") {
        count++;
      }
    });
    counts.push({ domain, count });
  });

  if (!counts.length) {
    if (bubbleChartInstance) bubbleChartInstance.destroy();
    return;
  }

  const maxCount = counts.reduce((m, c) => Math.max(m, c.count), 0) || 1;

  const palette = [
    "#2ecc71",
    "#3498db",
    "#9b59b6",
    "#e67e22",
    "#e74c3c",
    "#16a085",
    "#f1c40f",
    "#34495e"
  ];

  const spacing = Math.min(12, 80 / Math.max(counts.length - 1, 1));
  const startX = 10;

  const datasets = counts.map((item, i) => {
    const radius = 10 + (item.count / maxCount) * 25; // rayon entre 10 et 35
    const x = startX + i * spacing;
    const y = 50 + (i % 2 === 0 ? -6 : 6);

    const color = palette[i % palette.length];

    return {
      label: item.domain.label + " (" + item.count + ")",
      data: [{ x, y, r: radius, count: item.count }],
      backgroundColor: color,
      _count: item.count
    };
  });

  if (bubbleChartInstance) {
    bubbleChartInstance.destroy();
  }

  bubbleChartInstance = new Chart(ctx, {
    type: "bubble",
    plugins: [chartValueLabelPlugin],
    data: {
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
          labels: {
            boxWidth: 12,
            padding: 10
          }
        },
        tooltip: {
          callbacks: {
            label: (context) => context.dataset.label
          }
        },
        chartValueLabel: {
          bubbleColor: "#ffffff"
        }
      },
      scales: {
        x: {
          display: false,
          min: 0,
          max: 100
        },
        y: {
          display: false,
          min: 0,
          max: 100
        }
      }
    }
  });
}
const viewModeRadios = document.querySelectorAll('input[name="viewMode"]');
viewModeRadios.forEach(radio => {
  radio.addEventListener("change", (e) => {
    VIEW_MODE = e.target.value;
    if (VIEW_MODE === "memoire") {
      document.body.classList.add("memoire-mode");
      document.body.classList.add("layout-list");
      document.body.classList.remove("layout-cards");
      document.body.classList.remove("layout-bubbles");
      document.body.classList.remove("layout-pie");
      const listBtn = document.getElementById("layoutListBtn");
      const cardsBtn = document.getElementById("layoutCardsBtn");
      const pieBtn = document.getElementById("layoutPieBtn");
      const bubbleBtn = document.getElementById("layoutBubblesBtn");
      if (listBtn && cardsBtn) {
        listBtn.classList.add("active");
        cardsBtn.classList.remove("active");
        if (pieBtn) pieBtn.classList.remove("active");
        if (bubbleBtn) bubbleBtn.classList.remove("active");
      }
    } else {
      document.body.classList.remove("memoire-mode");
      document.body.classList.remove("layout-list");
      document.body.classList.remove("layout-cards");
      document.body.classList.remove("layout-bubbles");
      document.body.classList.remove("layout-pie");
    }
    applyFilters();
  });
});


let chartInstance = null;
let bubbleChartInstance = null;
let pieChartInstance = null;
let VIEW_MODE = "advanced";
let ALL_OPERATORS = [];      // { name, row, domains:{ key:{status,org,num,deb,fin} }, pole, section, manager, email }
let currentStatsGlobal = null; // stats par domaine pour tout le fichier
let LOCATION_MAP = {};       // key -> { email, section, pole, manager }
let HAS_LOCATION = false;
let LAST_FILTERED_OPERATORS = [];

document.getElementById("openFileBtn").addEventListener("click", () => {
  fileInput.click();
});
document.getElementById("openLocationBtn").addEventListener("click", () => {
  locationFileInput.click();
});

/* ==========
   Chargement fichier opérateurs
   ========== */

fileInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  showLoading();

  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, {type:"binary", cellDates:true});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {header:1});
    processRows(rows);
    hideLoading();
  };
  reader.readAsBinaryString(file);
});

/* ==========
   Chargement fichier géographique
   ========== */

locationFileInput.addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, {type:"binary", cellDates:true});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, {header:1});
    buildLocationMap(rows);
    HAS_LOCATION = true;

    if (ALL_OPERATORS.length) {
      attachLocationToAllOperators();
      populateLocationFilters(ALL_OPERATORS);
      applyFilters();
    }
  };
  reader.readAsBinaryString(file);
});

function buildLocationMap(rows) {
  LOCATION_MAP = {};

  rows.slice(1).forEach(row => {
    const nom = row[1];
    const prenom = row[2];
    if (!nom && !prenom) return;

    const email = row[3] || "";
    const section = row[4] || "—";
    const pole = row[5] || "—";
    const managerNom = row[6] || "";
    const managerPrenom = row[7] || "";
    const manager = (managerNom + " " + managerPrenom).trim() || "—";

    const key = getGeoKeyFromRow(row);
    LOCATION_MAP[key] = { email, section, pole, manager };
  });
}

/* ==========
   Barre de chargement
   ========== */

function showLoading(){
  loadingBarWrapper.style.display = "block";
  loadingBar.style.width = "0%";
  setTimeout(()=>loadingBar.style.width="60%",100);
}
function hideLoading(){
  loadingBar.style.width = "100%";
  setTimeout(()=>loadingBarWrapper.style.display="none",500);
}

/* ==========
   Traitement des lignes opérateurs
   ========== */

function processRows(rows){
  ALL_OPERATORS = [];
  tbody.innerHTML = "";

  let totalOps = 0;
  let expiredOps = 0;
  let fullCoreOps = 0;

  const stats = {};
  DOMAINS.forEach(d => stats[d.label] = {valid:0, expired:0, none:0});

  const certSet = new Set();

  rows.slice(1).forEach(row => {
    const name = (row[COL.NAME] || "")
      .toString()
      .trim()
      .replace(/\s+/g, " ");

    if (!name) return;

    totalOps++;

    let opExpired = false;
    let opFullCore = true;

    const op = {
      name,
      row,
      domains: {},
      pole: "—",
      section: "—",
      manager: "—",
      email: ""
    };

    DOMAINS.forEach(domain => {
      let status = "none";
      let org = "—";
      let num = "—";
      let deb = null;
      let fin = null;

      if (domain.audit != null) {
        const flag = row[domain.audit];
        if (flag) {
          status = "valid";
          org = flag.toString().trim() || "OK";
          stats[domain.label].valid++;
        } else {
          status = "none";
          stats[domain.label].none++;
        }
      } else {
        org = (row[domain.cert] || "").toString().trim() || "—";
        num = (row[domain.num] || "").toString().trim() || "—";
        deb = parseExcelDate(row[domain.deb]);
        fin = parseExcelDate(row[domain.fin]);

        const hasData = (org !== "—" || num !== "—" || deb || fin);

        if (!hasData) {
          status = "none";
          stats[domain.label].none++;
          if (domain.isCore) opFullCore = false;
        } else if (fin && !isValidDate(fin)) {
          status = "expired";
          stats[domain.label].expired++;
          opExpired = true;
          if (domain.isCore) opFullCore = false;
        } else {
          status = "valid";
          stats[domain.label].valid++;
        }
      }

      if (org && org !== "—") {
        certSet.add(org);
      }

      op.domains[domain.key] = {status, org, num, deb, fin};
    });

    if (opExpired) expiredOps++;
    if (opFullCore) fullCoreOps++;

    op.isFullCore = opFullCore;
    ALL_OPERATORS.push(op);
  });

  // KPI globaux
  kpiTotal.textContent = totalOps;
  kpiExpired.textContent = expiredOps;
  kpiFull.textContent = fullCoreOps;

  currentStatsGlobal = stats;

  populateCertFilter(certSet);
  populateOperatorSearch(ALL_OPERATORS);

  if (HAS_LOCATION) {
    attachLocationToAllOperators();
    populateLocationFilters(ALL_OPERATORS);
  }

  drawChart();
  applyFilters();
}

function attachLocationToAllOperators() {
  ALL_OPERATORS.forEach(op => {
    const key = getGeoKeyFromOperator(op);
    const loc = LOCATION_MAP[key];

    if (loc) {
      op.email = loc.email || "";
      op.section = loc.section || "—";
      op.pole = loc.pole || "—";
      op.manager = loc.manager || "—";
    } else {
      op.email = "";
      op.section = "—";
      op.pole = "—";
      op.manager = "—";
    }
  });
}

function populateLocationFilters(ops) {
  const poles = new Set();
  const sections = new Set();

  ops.forEach(op => {
    if (op.pole && op.pole !== "—") poles.add(op.pole);
    if (op.section && op.section !== "—") sections.add(op.section);
  });

  filterPoleSelect.innerHTML = '<option value="ALL">Tous</option>';
  Array.from(poles).sort().forEach(p => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    filterPoleSelect.appendChild(opt);
  });

  filterSectionSelect.innerHTML = '<option value="ALL">Toutes</option>';
  Array.from(sections).sort().forEach(s => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    filterSectionSelect.appendChild(opt);
  });
}

function populateOperatorSearch(ops) {
  const select = searchOperatorSelect;
  select.innerHTML = '<option value="ALL">Tous</option>';

  const sortedNames = ops.map(o => o.name).sort((a, b) => a.localeCompare(b));

  sortedNames.forEach(name => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

/* ==========
   Construction du filtre certificateur
   ========== */

function populateCertFilter(certSet){
  filterCertSelect.innerHTML = '<option value="ALL">Tous</option>';

  const certs = Array.from(certSet).sort((a,b)=>a.localeCompare(b));
  certs.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    filterCertSelect.appendChild(opt);
  });
}

/* ==========
   Rendu du tableau
   ========== */

function renderTable(operators){
  tbody.innerHTML = "";

  operators.forEach(op => {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td>' + op.name + '</td>';

    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key] || {
        status:"none", org:"—", num:"—", deb:null, fin:null
      };

      const dDeb = d.deb ? d.deb.toLocaleDateString() : "—";
      const dFin = d.fin ? d.fin.toLocaleDateString() : "—";

      let tooltip = domain.label + "\n"
        + "Certificateur : " + d.org + "\n"
        + "Numéro : " + d.num + "\n"
        + "Début : " + dDeb + "\n"
        + "Fin : " + dFin;

      if (domain.key === "audit") {
        tooltip = (d.status === "valid")
          ? "Audit : Certifié (AW rempli)"
          : "Audit : Non certifié";
      }

      let extraClass = "";
      if (d.status === "valid" && d.fin) {
        const today = new Date();
        today.setHours(0,0,0,0);
        const finDate = new Date(d.fin.getTime());
        finDate.setHours(0,0,0,0);
        const diffMs = finDate - today;
        const diffDays = diffMs / (1000*60*60*24);
        const diffMonths = diffDays / 30;

        if (diffMonths <= 6) {
          extraClass = " near6";
        } else if (diffMonths <= 12) {
          extraClass = " near12";
        }
      }

      const td = document.createElement("td");
      const isMemoire = (typeof VIEW_MODE !== "undefined" && VIEW_MODE === "memoire");

      if (isMemoire && (d.status === "expired" || d.status === "none")) {
        td.textContent = "—";
        tr.appendChild(td);
        return;
      }

      const safeOrg = d.org === "" ? "—" : d.org;
      const safeTooltip = tooltip.replace(/"/g, '&quot;');

      td.innerHTML =
        '<div class="cell-content">'
        +   '<span class="dot ' + d.status + extraClass + '" title="' + safeTooltip + '"></span>'
        +   '<span class="cert-org ' + d.status + '" title="' + safeTooltip + '" data-cert="' + safeOrg + '">' + safeOrg + '</span>'
        + '</div>';
      tr.appendChild(td);
    });

    const tdPole = document.createElement("td");
    tdPole.className = "geo-col";
    tdPole.textContent = op.pole || "—";
    tr.appendChild(tdPole);

    const tdSection = document.createElement("td");
    tdSection.className = "geo-col";
    tdSection.textContent = op.section || "—";
    tr.appendChild(tdSection);

    const tdManager = document.createElement("td");
    tdManager.className = "geo-col";
    tdManager.textContent = op.manager || "—";
    tr.appendChild(tdManager);

    tbody.appendChild(tr);
  });

  document.querySelectorAll(".cert-org").forEach(span => {
    const cert = span.getAttribute("data-cert");
    if (cert && cert !== "—") {
      span.style.cursor = "pointer";
      span.addEventListener("click", () => {
        filterCertSelect.value = cert;
        applyFilters();
      });
    }
  });
}

/* ==========
   Stats par certificateur / pôle / section
   ========== */

function computeStatsByCertificateur(){
  const stats = {};

  ALL_OPERATORS.forEach(op => {
    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key];
      if (!d) return;
      const org = d.org;
      if (!org || org === "—") return;

      if (!stats[org]) {
        stats[org] = {valid:0, expired:0, none:0};
      }
      stats[org][d.status]++;
    });
  });

  return stats;
}

function computeStatsByPole(){
  const stats = {};
  ALL_OPERATORS.forEach(op => {
    const pole = (op.pole && op.pole !== "—") ? op.pole : null;
    if (!pole) return;
    if (!stats[pole]) {
      stats[pole] = {valid:0, expired:0, none:0};
    }
    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key];
      if (!d) return;
      stats[pole][d.status]++;
    });
  });
  return stats;
}

function computeStatsBySection(){
  const stats = {};
  ALL_OPERATORS.forEach(op => {
    const section = (op.section && op.section !== "—") ? op.section : null;
    if (!section) return;
    if (!stats[section]) {
      stats[section] = {valid:0, expired:0, none:0};
    }
    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key];
      if (!d) return;
      stats[section][d.status]++;
    });
  });
  return stats;
}

/* ==========
   Graph global : domaines / cert / pôle / section
   ========== */

function drawChart(){
  if (!currentStatsGlobal) return;

  const mode = graphModeSelect.value;

  if (chartInstance) chartInstance.destroy();

  let labels = [];
  let validData = [];
  let expiredData = [];
  let noneData = [];

  if (mode === "domain") {
    labels = DOMAINS.map(d => d.label);
    validData   = labels.map(l => currentStatsGlobal[l].valid);
    expiredData = labels.map(l => currentStatsGlobal[l].expired);
    noneData    = labels.map(l => currentStatsGlobal[l].none);
  } else if (mode === "cert") {
    const certStats = computeStatsByCertificateur();
    labels = Object.keys(certStats).sort((a,b)=>a.localeCompare(b));
    validData   = labels.map(c => certStats[c].valid);
    expiredData = labels.map(c => certStats[c].expired);
    noneData    = labels.map(c => certStats[c].none);
  } else if (mode === "pole") {
    const poleStats = computeStatsByPole();
    labels = Object.keys(poleStats).sort((a,b)=>a.localeCompare(b));
    validData   = labels.map(p => poleStats[p].valid);
    expiredData = labels.map(p => poleStats[p].expired);
    noneData    = labels.map(p => poleStats[p].none);
  } else if (mode === "section") {
    const sectionStats = computeStatsBySection();
    labels = Object.keys(sectionStats).sort((a,b)=>a.localeCompare(b));
    validData   = labels.map(s => sectionStats[s].valid);
    expiredData = labels.map(s => sectionStats[s].expired);
    noneData    = labels.map(s => sectionStats[s].none);
  }

  chartInstance = new Chart(chartCanvas.getContext("2d"), {
    type:"bar",
    data:{
      labels,
      datasets:[
        {label:"Valide", data:validData, backgroundColor:"#2ecc71"},
        {label:"Expirée", data:expiredData, backgroundColor:"#e74c3c"},
        {label:"Aucune", data:noneData, backgroundColor:"#bdc3c7"}
      ]
    },
    options:{
      responsive:true,
      plugins:{ legend:{ position:"bottom" } },
      scales:{ x:{ stacked:true }, y:{ stacked:true } }
    }
  });
}

/* ==========
   Pastilles en tête de colonnes (valid / expiré / 6 / 12 mois)
   ========== */


function updateHeaderBadges(ops) {
  const counters = {};
  DOMAINS.forEach(d => {
    counters[d.key] = {
      valid: 0,
      expired: 0,
      six: 0,
      twelve: 0
    };
  });

  const today = new Date();
  today.setHours(0,0,0,0);

  const plus6  = new Date(today.getFullYear(), today.getMonth()+6,  today.getDate());
  const plus12 = new Date(today.getFullYear(), today.getMonth()+12, today.getDate());

  ops.forEach(op => {
    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key];
      if (!d) return;

      if (d.status === "expired") {
        counters[domain.key].expired++;
        return;
      }

      if (d.status === "valid") {
        if (!d.fin) {
          counters[domain.key].valid++;
          return;
        }

        const fin = new Date(d.fin);
        fin.setHours(0,0,0,0);

        if (fin <= plus6) counters[domain.key].six++;
        else if (fin <= plus12) counters[domain.key].twelve++;
        else counters[domain.key].valid++;
      }
    });
  });

  DOMAINS.forEach(domain => {
    const slot = document.getElementById("H_" + domain.key);
    const c = counters[domain.key];
    const isMemoire = (typeof VIEW_MODE !== "undefined" && VIEW_MODE === "memoire");
    let html = "";

    if (isMemoire) {
      // En mode mémoires : une seule pastille verte = toutes les certifs valides (y compris proches)
      const totalValid = c.valid + c.six + c.twelve;
      if (totalValid > 0) {
        html += '<span class="h-badge h-green" title="Certifications valides">' + totalValid + '</span>';
      }
    } else {
      // Mode avancé : on garde le détail (vert / rouge / violet / orange)
      if (c.valid > 0)
        html += '<span class="h-badge h-green" title="Valides">' + c.valid + '</span>';
      if (c.expired > 0)
        html += '<span class="h-badge h-red" title="Expirées">' + c.expired + '</span>';
      if (c.six > 0)
        html += '<span class="h-badge h-purple" title="≤ 6 mois">' + c.six + '</span>';
      if (c.twelve > 0)
        html += '<span class="h-badge h-orange" title="≤ 12 mois">' + c.twelve + '</span>';
    }

    slot.innerHTML = html;
  });
}


/* ==========
   Application combinée des filtres
   ========== */

function applyFilters(){
  if (!ALL_OPERATORS.length) return;

  const months = parseInt(filterSelect.value, 10);
  const certVal = filterCertSelect.value;
  const searchText = searchOperatorInput.value.trim().toLowerCase();
  const poleVal = filterPoleSelect.value;
  const sectionVal = filterSectionSelect.value;

  let ops = ALL_OPERATORS.slice();

  const isMemoire = (VIEW_MODE === "memoire");

  if (searchText !== "") {
    ops = ops.filter(op => op.name.toLowerCase().includes(searchText));
  }

  if (!isMemoire && months > 0) {
    const today = new Date();
    today.setHours(0,0,0,0);
    const target = new Date(today.getTime());
    target.setMonth(target.getMonth() + months);

    ops = ops.filter(op => {
      return Object.values(op.domains).some(d => {
        if (!d.fin) return false;
        const fin = d.fin;
        return fin >= today && fin <= target;
      });
    });
  }

  if (certVal !== "ALL") {
    ops = ops.filter(op => {
      return Object.values(op.domains).some(d => d.org === certVal);
    });
  }

  if (poleVal !== "ALL") {
    ops = ops.filter(op => op.pole === poleVal);
  }

  if (sectionVal !== "ALL") {
    ops = ops.filter(op => op.section === sectionVal);
  }

  document.getElementById("kpiFilteredTotal").textContent = ops.length;

  const expiredFiltered = ops.filter(op =>
    Object.values(op.domains).some(d => d.status === "expired")
  ).length;
  document.getElementById("kpiFilteredExpired").textContent = expiredFiltered;

  const fullFiltered = ops.filter(op =>
    DOMAINS.filter(d => d.isCore).every(d => op.domains[d.key].status === "valid")
  ).length;
  document.getElementById("kpiFilteredFull").textContent = fullFiltered;

  LAST_FILTERED_OPERATORS = ops;

  renderTable(ops);
  updateHeaderBadges(ops);
  updateDomainVisibility();
  renderMemoirePanels(ops);

  if (document.body.classList.contains("layout-pie")) {
    updatePieChart(ops);
  }
  if (document.body.classList.contains("layout-bubbles")) {
    updateBubbleChart(ops);
  }
}

/* ==========
   Export tableau
   ========== */

function buildOperatorExportTable(operators){
  let html = ''
    + '<table border="1" cellpadding="5" style="border-collapse:collapse;width:100%;font-size:12px;">'
    + '<thead>'
    + '<tr>'
    + '<th>Opérateur</th>'
    + '<th>Amiante</th>'
    + '<th>CREP</th>'
    + '<th>Termites</th>'
    + '<th>DPE Mention</th>'
    + '<th>Gaz</th>'
    + '<th>Élec</th>'
    + '<th>DPE Indiv</th>'
    + '<th>Audit</th>'
    + '<th>Pôle</th>'
    + '<th>Section</th>'
    + '<th>Manager</th>'
    + '</tr>'
    + '</thead>'
    + '<tbody>';

  operators.forEach(op => {
    html += '<tr><td><b>' + op.name + '</b></td>';

    DOMAINS.forEach(domain => {
      const d = op.domains[domain.key];
      if (!d) {
        html += "<td>—</td>";
        return;
      }

      const dDeb = d.deb ? d.deb.toLocaleDateString() : "—";
      const dFin = d.fin ? d.fin.toLocaleDateString() : "—";

      html += '<td>'
        + '<b>' + d.status.toUpperCase() + '</b><br>'
        + 'Certif : ' + (d.org || "—") + '<br>'
        + 'Num : ' + (d.num || "—") + '<br>'
        + 'Début : ' + dDeb + '<br>'
        + 'Fin : ' + dFin
        + '</td>';
    });

    html += '<td>' + (op.pole || "—") + '</td>'
         +  '<td>' + (op.section || "—") + '</td>'
         +  '<td>' + (op.manager || "—") + '</td>'
         + '</tr>';
  });

  html += '</tbody></table>';
  return html;
}

function openExportVisual(){
  const rows = Array.from(document.querySelectorAll("#dataTable tbody tr"));
  if (rows.length === 0){
    alert("Aucune donnée à exporter.");
    return;
  }

  const ops = [];
  rows.forEach(tr => {
    const name = tr.children[0].innerText.trim();
    const op = ALL_OPERATORS.find(o => o.name === name);
    if (op) ops.push(op);
  });

  const tableHTML = buildOperatorExportTable(ops);
  exportTableContainer.innerHTML = tableHTML;

  exportModal.style.display = "flex";
}

function copyExportTable(){
  const htmlTable = exportTableContainer.innerHTML;

  if (!navigator.clipboard) {
    alert("Votre navigateur ne supporte pas la copie HTML.");
    return;
  }

  const blob = new Blob([htmlTable], { type: "text/html" });

  navigator.clipboard.write([
    new ClipboardItem({
      "text/html": blob
    })
  ])
  .then(() => {
    alert("Tableau copié avec mise en forme ! Vous pouvez coller dans Excel pour exploiter les données.");
  })
  .catch(err => {
    alert("Erreur de copie. Sélectionnez et copiez manuellement.");
    console.error(err);
  });
}

/* ==========
   Export emails
   ========== */

function exportEmails() {
  const rows = Array.from(document.querySelectorAll("#dataTable tbody tr"));
  if (rows.length === 0) {
    alert("Aucune donnée à exporter.");
    return;
  }

  const emailSet = new Set();

  rows.forEach(tr => {
    const name = tr.children[0].innerText.trim();
    const op = ALL_OPERATORS.find(o => o.name === name);
    if (op && op.email) {
      emailSet.add(op.email.trim());
    }
  });

  if (emailSet.size === 0) {
    alert("Aucun email disponible pour les opérateurs filtrés.");
    return;
  }

  const list = Array.from(emailSet).join(";");
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(list)
      .then(() => {
        alert("Liste d\'emails copiée dans le presse-papiers :\n" + list);
      })
      .catch(() => {
        alert("Voici la liste des emails :\n" + list);
      });
  } else {
    alert("Voici la liste des emails :\n" + list);
  }
}

/* ==========
   Toggle affichage des certificateurs
   ========== */

document.body.classList.add("hide-cert");

toggleCert.addEventListener("change", () => {
  if (toggleCert.checked) {
    document.body.classList.remove("hide-cert");
  } else {
    document.body.classList.add("hide-cert");
  }
});

/* ==========
   Écouteurs filtres & mode graphique & export
   ========== */

filterSelect.addEventListener("change", applyFilters);
filterCertSelect.addEventListener("change", applyFilters);
filterPoleSelect.addEventListener("change", applyFilters);
filterSectionSelect.addEventListener("change", applyFilters);

graphModeSelect.addEventListener("change", () => {
  if (!currentStatsGlobal) return;
  drawChart();
});

document.getElementById("toggleChart").addEventListener("change", (e) => {
  const chartContainer = document.getElementById("chartContainer");
  if (e.target.checked) {
    chartContainer.style.display = "block";
    if (currentStatsGlobal) drawChart();
  } else {
    chartContainer.style.display = "none";
  }
});

toggleGeoColsCheckbox.addEventListener("change", (e) => {
  const table = document.getElementById("dataTable");
  if (e.target.checked) {
    table.classList.add("show-geo");
  } else {
    table.classList.remove("show-geo");
  }
});

exportVisualBtn.addEventListener("click", openExportVisual);
copyExportBtn.addEventListener("click", copyExportTable);
closeExportBtn.addEventListener("click", () => {
  exportModal.style.display = "none";
});
exportEmailsBtn.addEventListener("click", exportEmails);

searchOperatorInput.addEventListener("input", (e) => {
  searchOperatorSelect.value = "ALL";
  applyFilters();
});

searchOperatorSelect.addEventListener("change", (e) => {
  const val = e.target.value;
  if (val === "ALL") {
    searchOperatorInput.value = "";
  } else {
    searchOperatorInput.value = val;
  }
  applyFilters();
});

/* ==========
   Tri par domaine (clic sur l'en-tête)
   ========== */

const sortOrder = { expired:0, valid:1, none:2 };

function sortByDomain(domainKey){
  const domainIndex = DOMAINS.findIndex(d => d.key === domainKey);
  if (domainIndex < 0) return;

  const tdIndex = domainIndex + 1;

  const rows = Array.from(tbody.querySelectorAll("tr"));

  rows.sort((a,b) => {
    const aDot = a.children[tdIndex].querySelector(".dot");
    const bDot = b.children[tdIndex].querySelector(".dot");

    const aStatus = aDot ? aDot.classList[1] : "none";
    const bStatus = bDot ? bDot.classList[1] : "none";

    if (sortOrder[aStatus] !== sortOrder[bStatus]) {
      return sortOrder[aStatus] - sortOrder[bStatus];
    }

    const nameA = a.children[0].innerText.toLowerCase();
    const nameB = b.children[0].innerText.toLowerCase();
    return nameA.localeCompare(nameB);
  });

  tbody.innerHTML = "";
  rows.forEach(r => tbody.appendChild(r));
}

document.querySelectorAll("th[data-domain]").forEach(th => {
  th.addEventListener("click", () => {
    const key = th.getAttribute("data-domain");
    sortByDomain(key);
  });
  /* ==========
   Tri par colonnes géographiques
   ========== */

const geoSortState = { 9:1, 10:1, 11:1 };  // 1 = asc, -1 = desc

function sortGeoColumn(colIndex) {
  const rows = Array.from(tbody.querySelectorAll("tr"));

  rows.sort((a, b) => {
    const aText = a.children[colIndex].innerText.trim().toLowerCase();
    const bText = b.children[colIndex].innerText.trim().toLowerCase();

    return geoSortState[colIndex] * aText.localeCompare(bText);
  });

  geoSortState[colIndex] *= -1; // Inversion pour clic suivant

  tbody.innerHTML = "";
  rows.forEach(r => tbody.appendChild(r));
}

// Activation du clic sur les têtes (Pôle, Section, Manager)
document.querySelectorAll("#dataTable thead th.geo-col")
  .forEach((th, i) => {
    const colIndex = 9 + i; // indices 9,10,11
    th.style.cursor = "pointer";
    th.addEventListener("click", () => sortGeoColumn(colIndex));
  });

});

const layoutListBtn = document.getElementById("layoutListBtn");
const layoutCardsBtn = document.getElementById("layoutCardsBtn");
const layoutBubblesBtn = document.getElementById("layoutBubblesBtn");
const layoutPieBtn = document.getElementById("layoutPieBtn");

if (layoutListBtn && layoutCardsBtn && layoutBubblesBtn && layoutPieBtn) {
  layoutListBtn.addEventListener("click", () => {
    document.body.classList.add("layout-list");
    document.body.classList.remove("layout-cards");
    document.body.classList.remove("layout-bubbles");
    document.body.classList.remove("layout-pie");
    layoutListBtn.classList.add("active");
    layoutCardsBtn.classList.remove("active");
    layoutBubblesBtn.classList.remove("active");
    layoutPieBtn.classList.remove("active");
  });

  layoutCardsBtn.addEventListener("click", () => {
    document.body.classList.remove("layout-list");
    document.body.classList.add("layout-cards");
    document.body.classList.remove("layout-bubbles");
    document.body.classList.remove("layout-pie");
    layoutCardsBtn.classList.add("active");
    layoutListBtn.classList.remove("active");
    layoutBubblesBtn.classList.remove("active");
    layoutPieBtn.classList.remove("active");
  });

  layoutBubblesBtn.addEventListener("click", () => {
    document.body.classList.remove("layout-list");
    document.body.classList.remove("layout-cards");
    document.body.classList.add("layout-bubbles");
    document.body.classList.remove("layout-pie");
    layoutBubblesBtn.classList.add("active");
    layoutListBtn.classList.remove("active");
    layoutCardsBtn.classList.remove("active");
    layoutPieBtn.classList.remove("active");
    updateBubbleChart(LAST_FILTERED_OPERATORS);
  });

  layoutPieBtn.addEventListener("click", () => {
    document.body.classList.remove("layout-list");
    document.body.classList.remove("layout-cards");
    document.body.classList.remove("layout-bubbles");
    document.body.classList.add("layout-pie");
    layoutPieBtn.classList.add("active");
    layoutListBtn.classList.remove("active");
    layoutCardsBtn.classList.remove("active");
    layoutBubblesBtn.classList.remove("active");
    updatePieChart(LAST_FILTERED_OPERATORS);
  });
}
