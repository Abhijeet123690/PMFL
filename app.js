/* ============================================================
   PO Tracking Dashboard — Client-side logic
   ============================================================ */

let rawRows = [];          // all parsed rows
let currentCategory = null;
let currentCity = null;
let chartInstance = null;
let sortColumn = null;
let sortAsc = true;

const REQUIRED_COLUMNS = [
  'Category', 'City', 'Vendor_Name', 'Delivery_Location',
  'PO_No', 'PO_Date', 'PO_Expiry', 'Status', 'PO_QTY'
];

/* ---------- File Upload & Parsing ---------- */

document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;
  parseFile(file);
  e.target.value = ''; // allow re-uploading same file
});

function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  setStatus('Parsing…', '');

  if (ext === 'csv') {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: 'greedy',
      dynamicTyping: true,
      complete: (results) => handleParsed(results.data, results.meta.fields),
      error: (err) => setStatus('CSV parse error: ' + err.message, 'error')
    });
  } else if (ext === 'xlsx' || ext === 'xls') {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);
        const fields = data.length ? Object.keys(data[0]) : [];
        handleParsed(data, fields);
      } catch (err) {
        setStatus('Excel parse error: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  } else if (ext === 'json') {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result);
        const rows = Array.isArray(data) ? data : (data.rows || data.data || []);
        const fields = rows.length ? Object.keys(rows[0]) : [];
        handleParsed(rows, fields);
      } catch (err) {
        setStatus('JSON parse error: ' + err.message, 'error');
      }
    };
    reader.readAsText(file);
  } else {
    setStatus('Unsupported file type. Use .csv, .xlsx, .xls, or .json', 'error');
  }
}

function handleParsed(rows, fields) {
  // Normalize column names (trim)
  const normalized = rows.map(row => {
    const r = {};
    Object.keys(row).forEach(k => { r[k.trim()] = row[k]; });
    return r;
  });

  // Check required columns
  const missing = REQUIRED_COLUMNS.filter(c => !fields.includes(c));
  if (missing.length > 0) {
    setStatus('Missing columns: ' + missing.join(', '), 'error');
    return;
  }

  // Normalize and validate
  rawRows = normalized.map(r => ({
    Category: String(r.Category || '').trim(),
    City: String(r.City || '').trim(),
    Vendor_Name: String(r.Vendor_Name || '').trim(),
    Delivery_Location: String(r.Delivery_Location || '').trim(),
    PO_No: String(r.PO_No || '').trim(),
    PO_Date: normalizeDate(r.PO_Date),
    PO_Expiry: normalizeDate(r.PO_Expiry),
    Appointment_Date: normalizeDate(r.Appointment_Date),
    Status: String(r.Status || '').trim(),
    PO_QTY: Number(r.PO_QTY) || 0
  })).filter(r => r.PO_No && r.Category && r.City);

  if (rawRows.length === 0) {
    setStatus('No valid rows found. Check that PO_No, Category, and City are populated.', 'error');
    return;
  }

  setStatus(`Loaded ${rawRows.length} rows.`, 'success');
  resetToPhase1();
}

function normalizeDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  const s = String(val).trim();
  // Try to parse and reformat
  const d = new Date(s);
  if (!isNaN(d.getTime()) && s.match(/\d/)) {
    return d.toISOString().slice(0, 10);
  }
  return s;
}

function setStatus(msg, cls) {
  const el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = 'status-msg ' + (cls || '');
}

/* ---------- Aggregation Helpers ---------- */

function groupBy(rows, keyFn) {
  const map = new Map();
  rows.forEach(r => {
    const k = keyFn(r);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  });
  return map;
}

function distinctPOs(rows) {
  return new Set(rows.map(r => r.PO_No)).size;
}

function sumQTY(rows) {
  return rows.reduce((s, r) => s + r.PO_QTY, 0);
}

function countStatus(rows, status) {
  return rows.filter(r => r.Status === status).length;
}

function daysUntil(dateStr) {
  if (!dateStr) return Infinity;
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return Infinity;
  return Math.ceil((d - new Date()) / (1000 * 60 * 60 * 24));
}

/* ---------- Phase 1 ---------- */

function resetToPhase1() {
  currentCategory = null;
  currentCity = null;
  document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
  document.getElementById('phase1').classList.add('active');
  renderBreadcrumb();
  renderPhase1();
}

function renderPhase1() {
  renderKPIs();
  renderCategoryChart();
  renderCategoryTable();
}

function renderKPIs() {
  const totalPOs = distinctPOs(rawRows);
  const totalQTY = sumQTY(rawRows);
  const open = countStatus(rawRows, 'Open');
  const expired = countStatus(rawRows, 'Expired');
  const expiring30 = rawRows.filter(r => {
    const d = daysUntil(r.PO_Expiry);
    return d >= 0 && d <= 30 && r.Status !== 'Delivered' && r.Status !== 'Cancelled';
  }).length;

  document.getElementById('kpiCards').innerHTML = `
    <div class="kpi-card"><div class="kpi-label">Total POs</div><div class="kpi-value">${totalPOs.toLocaleString()}</div></div>
    <div class="kpi-card"><div class="kpi-label">Total PO QTY</div><div class="kpi-value">${totalQTY.toLocaleString()}</div></div>
    <div class="kpi-card"><div class="kpi-label">Open POs</div><div class="kpi-value">${open.toLocaleString()}</div></div>
    <div class="kpi-card"><div class="kpi-label">Expired</div><div class="kpi-value">${expired.toLocaleString()}</div></div>
    <div class="kpi-card"><div class="kpi-label">Expiring ≤30d</div><div class="kpi-value">${expiring30.toLocaleString()}</div></div>
  `;
}

function renderCategoryChart() {
  const catMap = groupBy(rawRows, r => r.Category);
  const labels = [];
  const values = [];
  catMap.forEach((rows, cat) => {
    labels.push(cat);
    values.push(sumQTY(rows));
  });

  // Sort descending
  const order = labels.map((l, i) => i).sort((a, b) => values[b] - values[a]);
  const sortedLabels = order.map(i => labels[i]);
  const sortedValues = order.map(i => values[i]);

  const ctx = document.getElementById('categoryChart').getContext('2d');
  if (chartInstance) chartInstance.destroy();
  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sortedLabels,
      datasets: [{
        label: 'Sum of PO QTY',
        data: sortedValues,
        backgroundColor: 'rgba(37, 99, 235, 0.7)',
        borderColor: 'rgba(37, 99, 235, 1)',
        borderWidth: 1
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true } }
    }
  });
}

function renderCategoryTable() {
  const catMap = groupBy(rawRows, r => r.Category);
  const tbody = document.querySelector('#categoryTable tbody');
  tbody.innerHTML = '';

  const entries = [];
  catMap.forEach((rows, cat) => {
    entries.push({
      category: cat,
      pos: distinctPOs(rows),
      qty: sumQTY(rows),
      open: countStatus(rows, 'Open'),
      delivered: countStatus(rows, 'Delivered'),
      expired: countStatus(rows, 'Expired'),
      expiring30: rows.filter(r => {
        const d = daysUntil(r.PO_Expiry);
        return d >= 0 && d <= 30 && r.Status !== 'Delivered' && r.Status !== 'Cancelled';
      }).length
    });
  });

  entries.sort((a, b) => b.qty - a.qty);

  entries.forEach(e => {
    const pct = e.pos > 0 ? ((e.expiring30 / e.pos) * 100).toFixed(1) : '0.0';
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${e.category}</strong></td>
      <td>${e.pos.toLocaleString()}</td>
      <td>${e.qty.toLocaleString()}</td>
      <td>${e.open.toLocaleString()}</td>
      <td>${e.delivered.toLocaleString()}</td>
      <td>${e.expired.toLocaleString()}</td>
      <td>${pct}%</td>
    `;
    tr.addEventListener('click', () => goToPhase2(e.category));
    tbody.appendChild(tr);
  });
}

/* ---------- Phase 2 ---------- */

function goToPhase2(category) {
  currentCategory = category;
  currentCity = null;
  document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
  document.getElementById('phase2').classList.add('active');
  renderBreadcrumb();
  renderPhase2();
}

function renderPhase2() {
  document.getElementById('phase2Title').textContent = `Category: ${currentCategory}`;
  const catRows = rawRows.filter(r => r.Category === currentCategory);
  const cityMap = groupBy(catRows, r => r.City);
  const tbody = document.querySelector('#cityTable tbody');
  tbody.innerHTML = '';

  const entries = [];
  cityMap.forEach((rows, city) => {
    const expiries = rows.map(r => r.PO_Expiry).filter(Boolean).sort();
    entries.push({
      city,
      pos: distinctPOs(rows),
      qty: sumQTY(rows),
      open: countStatus(rows, 'Open'),
      expired: countStatus(rows, 'Expired'),
      nearest: expiries[0] || '—'
    });
  });

  entries.sort((a, b) => b.qty - a.qty);

  entries.forEach(e => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${e.city}</strong></td>
      <td>${e.pos.toLocaleString()}</td>
      <td>${e.qty.toLocaleString()}</td>
      <td>${e.open.toLocaleString()}</td>
      <td>${e.expired.toLocaleString()}</td>
      <td>${e.nearest}</td>
    `;
    tr.addEventListener('click', () => goToPhase3(e.city));
    tbody.appendChild(tr);
  });
}

/* ---------- Phase 3 ---------- */

function goToPhase3(city) {
  currentCity = city;
  document.querySelectorAll('.phase').forEach(p => p.classList.remove('active'));
  document.getElementById('phase3').classList.add('active');
  renderBreadcrumb();
  renderPhase3();
}

function renderPhase3() {
  document.getElementById('phase3Title').textContent =
    `${currentCategory} › ${currentCity}`;
  renderDetailTable();
}

function renderDetailTable() {
  const filter = document.getElementById('statusFilter').value;
  let rows = rawRows.filter(r =>
    r.Category === currentCategory && r.City === currentCity
  );
  if (filter) rows = rows.filter(r => r.Status === filter);

  // Aggregate by PO_No (sum QTY per PO, keep first row's other fields)
  const poMap = new Map();
  rows.forEach(r => {
    if (!poMap.has(r.PO_No)) {
      poMap.set(r.PO_No, { ...r, PO_QTY: 0 });
    }
    poMap.get(r.PO_No).PO_QTY += r.PO_QTY;
  });

  let poRows = Array.from(poMap.values());

  // Sorting
  if (sortColumn) {
    poRows.sort((a, b) => {
      let va = a[sortColumn], vb = b[sortColumn];
      if (sortColumn === 'PO_QTY') { va = Number(va); vb = Number(vb); }
      if (va < vb) return sortAsc ? -1 : 1;
      if (va > vb) return sortAsc ? 1 : -1;
      return 0;
    });
  }

  const tbody = document.querySelector('#detailTable tbody');
  tbody.innerHTML = '';

  poRows.forEach(r => {
    const days = daysUntil(r.PO_Expiry);
    let rowClass = '';
    if (r.Status === 'Expired') rowClass = 'status-expired';
    else if (r.Status === 'Delivered') rowClass = 'status-delivered';
    else if (days >= 0 && days <= 14) rowClass = 'status-expiring';

    const tr = document.createElement('tr');
    if (rowClass) tr.className = rowClass;
    tr.innerHTML = `
      <td>${r.Vendor_Name}</td>
      <td>${r.Delivery_Location}</td>
      <td>${r.PO_No}</td>
      <td>${r.PO_Date || '—'}</td>
      <td>${r.PO_Expiry || '—'}</td>
      <td>${r.Appointment_Date || '—'}</td>
      <td>${r.Status}</td>
      <td>${r.PO_QTY.toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  });

  // Store current filtered rows for export
  tbody._currentRows = poRows;
}

/* ---------- Sorting ---------- */

document.querySelectorAll('#detailTable th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.sort;
    if (sortColumn === col) {
      sortAsc = !sortAsc;
    } else {
      sortColumn = col;
      sortAsc = true;
    }
    renderDetailTable();
  });
});

/* ---------- Status Filter ---------- */

document.getElementById('statusFilter').addEventListener('change', renderDetailTable);

/* ---------- Export CSV ---------- */

document.getElementById('exportBtn').addEventListener('click', () => {
  const tbody = document.querySelector('#detailTable tbody');
  const rows = tbody._currentRows || [];
  if (rows.length === 0) { alert('No data to export.'); return; }

  const headers = ['Vendor Name', 'Delivery Location', 'PO No', 'PO Date',
    'PO Expiry', 'Appointment Date', 'Status', 'Sum of PO QTY'];
  const csvRows = [headers.join(',')];
  rows.forEach(r => {
    csvRows.push([
      `"${r.Vendor_Name}"`,
      `"${r.Delivery_Location}"`,
      `"${r.PO_No}"`,
      r.PO_Date,
      r.PO_Expiry,
      r.Appointment_Date,
      `"${r.Status}"`,
      r.PO_QTY
    ].join(','));
  });

  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `po-export-${currentCategory}-${currentCity}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

/* ---------- Breadcrumb ---------- */

function renderBreadcrumb() {
  const bc = document.getElementById('breadcrumb');
  let html = '<a data-nav="home">All Categories</a>';
  if (currentCategory) {
    html += ` › <a data-nav="cat">${currentCategory}</a>`;
  }
  if (currentCity) {
    html += ` › <span>${currentCity}</span>`;
  }
  bc.innerHTML = html;

  bc.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', () => {
      const nav = a.dataset.nav;
      if (nav === 'home') resetToPhase1();
      else if (nav === 'cat') goToPhase2(currentCategory);
    });
  });
}

/* ---------- Init ---------- */

// Try to load a default dataset from data/latest.csv if present
fetch('data/latest.csv')
  .then(r => { if (!r.ok) throw new Error('no default'); return r.text(); })
  .then(text => {
    Papa.parse(text, {
      header: true, skipEmptyLines: 'greedy', dynamicTyping: true,
      complete: (res) => handleParsed(res.data, res.meta.fields)
    });
  })
  .catch(() => {
    setStatus('Upload a CSV, Excel, or JSON file to begin.', '');
  });
