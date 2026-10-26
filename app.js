let rawRows = [];
let choicesCategory, choicesCity, choicesStatus;

document.getElementById('currentDate').textContent = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' }).toUpperCase();

document.getElementById('fileInput').addEventListener('change', function (e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('fileName').textContent = file.name;
  parseFile(file);
});

function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  setStatus('Parsing…', '');

  if (ext === 'csv') {
    Papa.parse(file, {
      header: true, skipEmptyLines: 'greedy',
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
  }
}

// Normalize headers (ignores spaces, underscores, case)
function cleanKey(str) {
  return String(str).replace(/[\s_]/g, '').toLowerCase();
}

function handleParsed(rows, fields) {
  if (!rows || rows.length === 0) { setStatus('The uploaded file is empty.', 'error'); return; }

  const map = {};
  fields.forEach(f => { map[cleanKey(f)] = f; });

  // Check required columns based on your actual data
  const required = ['Category', 'City', 'Vendor Name', 'Del Location', 'PO No', 'PO Date', 'PO Expiry', 'Status', 'PO QTY', 'GRN QTY', 'Po Value', 'Po Volume In Kg', 'GRN Volume In Kg'];
  const missing = required.filter(req => !map[cleanKey(req)]);

  if (missing.length > 0) {
    setStatus('Missing columns: ' + missing.join(', ') + '. Please check headers.', 'error');
    return;
  }

  // Normalize and calculate
  rawRows = rows.map(r => {
    const getVal = (key) => r[map[cleanKey(key)]] !== undefined ? r[map[cleanKey(key)]] : '';
    const getNum = (key) => Number(getVal(key)) || 0;

    return {
      Category: String(getVal('Category')).trim(),
      City: String(getVal('City')).trim(),
      VendorName: String(getVal('Vendor Name')).trim(),
      DelLocation: String(getVal('Del Location')).trim(),
      PONo: String(getVal('PO No')).trim(),
      PODate: normalizeDate(getVal('PO Date')),
      POExpiry: normalizeDate(getVal('PO Expiry')),
      ApptDate: normalizeDate(getVal('Appointment Date')),
      Status: String(getVal('Status')).trim(),
      POQty: getNum('PO QTY'),
      GRNQty: getNum('GRN QTY'),
      PoValue: getNum('Po Value'),
      FillRate: getNum('Fill Rate'),
      Month: String(getVal('Month')).trim(),
      POTons: getNum('Po Volume In Kg') / 1000,
      GRNTons: getNum('GRN Volume In Kg') / 1000
    };
  }).filter(r => r.PONo && r.Category);

  setStatus(`Successfully loaded ${rawRows.length} rows.`, 'success');
  initializeDashboard();
}

function normalizeDate(val) {
  if (!val) return '';
  const s = String(val).trim();
  const d = new Date(s);
  if (!isNaN(d.getTime()) && s.match(/\d/)) {
    // If Excel serial date (just a number)
    if (/^\d+$/.test(s)) {
        const date = new Date((Number(s) - 25569) * 86400 * 1000);
        return date.toISOString().slice(0, 10);
    }
    return d.toISOString().slice(0, 10);
  }
  return s;
}

function setStatus(msg, cls) {
  const el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = 'status-msg ' + (cls || '');
}

// Initialize Dashboard & Multi-Selects
function initializeDashboard() {
  document.getElementById('filterSection').style.display = 'block';
  document.getElementById('kpiGrid').style.display = 'grid';
  document.getElementById('mainContent').style.display = 'block';

  const cats = [...new Set(rawRows.map(r => r.Category))].sort();
  const cities = [...new Set(rawRows.map(r => r.City))].sort();
  const statuses = [...new Set(rawRows.map(r => r.Status))].sort();

  // Initialize Choices.js Multi-selects
  const config = { removeItemButton: true, placeholder: true, placeholderValue: 'All', searchEnabled: true };
  
  if (choicesCategory) choicesCategory.destroy();
  if (choicesCity) choicesCity.destroy();
  if (choicesStatus) choicesStatus.destroy();

  choicesCategory = new Choices('#filterCategory', { ...config, choices: cats.map(c => ({ value: c, label: c })) });
  choicesCity = new Choices('#filterCity', { ...config, choices: cities.map(c => ({ value: c, label: c })) });
  choicesStatus = new Choices('#filterStatus', { ...config, choices: statuses.map(s => ({ value: s, label: s })) });

  // Add Event Listeners
  document.getElementById('filterCategory').addEventListener('change', updateDashboard);
  document.getElementById('filterCity').addEventListener('change', updateDashboard);
  document.getElementById('filterStatus').addEventListener('change', updateDashboard);
  document.getElementById('searchInput').addEventListener('input', updateDashboard);
  
  updateDashboard();
}

// Reset Filters
document.getElementById('resetBtn').addEventListener('click', () => {
  choicesCategory.removeActiveItems();
  choicesCity.removeActiveItems();
  choicesStatus.removeActiveItems();
  document.getElementById('searchInput').value = '';
  updateDashboard();
});

// Update Main Logic
function updateDashboard() {
  const selectedCats = choicesCategory.getValue(true);
  const selectedCities = choicesCity.getValue(true);
  const selectedStatuses = choicesStatus.getValue(true);
  const searchStr = document.getElementById('searchInput').value.toLowerCase();

  const filtered = rawRows.filter(r => {
    if (selectedCats.length > 0 && !selectedCats.includes(r.Category)) return false;
    if (selectedCities.length > 0 && !selectedCities.includes(r.City)) return false;
    if (selectedStatuses.length > 0 && !selectedStatuses.includes(r.Status)) return false;
    if (searchStr && !r.PONo.toLowerCase().includes(searchStr) && !r.VendorName.toLowerCase().includes(searchStr)) return false;
    return true;
  });

  document.getElementById('rowCount').textContent = `${filtered.length} filtered rows`;

  // Update KPIs
  const uniquePOs = new Set(filtered.map(r => r.PONo)).size;
  const totalQty = filtered.reduce((sum, r) => sum + r.POQty, 0);
  const openCount = filtered.filter(r => r.Status.toLowerCase() === 'open').length;
  const expiredCount = filtered.filter(r => r.Status.toLowerCase() === 'expired').length;
  const expiring = filtered.filter(r => {
    if (!r.POExpiry) return false;
    const days = Math.ceil((new Date(r.POExpiry) - new Date()) / (1000 * 60 * 60 * 24));
    return days >= 0 && days <= 30 && r.Status !== 'Delivered' && r.Status !== 'Cancelled';
  }).length;

  document.getElementById('kpiTotalPOs').textContent = uniquePOs.toLocaleString();
  document.getElementById('kpiTotalQty').textContent = totalQty.toLocaleString();
  document.getElementById('kpiOpen').textContent = openCount.toLocaleString();
  document.getElementById('kpiExpired').textContent = expiredCount.toLocaleString();
  document.getElementById('kpiExpiring').textContent = expiring.toLocaleString();

  // Determine Phase and Render Table
  if (selectedCats.length === 0 && selectedCities.length === 0) {
    document.getElementById('viewTitle').textContent = "Phase 1: Overall Category Level";
    renderPhase1(filtered);
  } else if (selectedCats.length === 1 && selectedCities.length === 0) {
    document.getElementById('viewTitle').textContent = `Phase 2: Category & City (${selectedCats[0]})`;
    renderPhase2(filtered);
  } else {
    document.getElementById('viewTitle').textContent = "Phase 3: Detailed PO View";
    renderPhase3(filtered);
  }
}

// PHASE 1: Category | PO | GRN
function renderPhase1(data) {
  const cols = ['Category', 'PO (Qty)', 'GRN (Qty)'];
  document.getElementById('tableHead').innerHTML = `<tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>`;
  
  const groups = {};
  data.forEach(r => {
    if (!groups[r.Category]) groups[r.Category] = { po: 0, grn: 0 };
    groups[r.Category].po += r.POQty;
    groups[r.Category].grn += r.GRNQty;
  });

  let html = '';
  Object.keys(groups).sort().forEach(cat => {
    html += `<tr>
      <td><strong>${cat}</strong></td>
      <td>${groups[cat].po.toLocaleString()}</td>
      <td>${groups[cat].grn.toLocaleString()}</td>
    </tr>`;
  });
  document.getElementById('tableBody').innerHTML = html || `<tr><td colspan="3" style="text-align:center;">No data</td></tr>`;
}

// PHASE 2: Category | City | PO Tons | GRN Tons | Fill Rate | Open Po | EXPIRED | CANCELLED
function renderPhase2(data) {
  const cols = ['Category', 'City', 'PO Tons', 'GRN Tons', 'Fill Rate %', 'Open Po', 'EXPIRED', 'CANCELLED'];
  document.getElementById('tableHead').innerHTML = `<tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>`;
  
  const groups = {};
  data.forEach(r => {
    const key = r.Category + '|' + r.City;
    if (!groups[key]) groups[key] = { cat: r.Category, city: r.City, poT: 0, grnT: 0, poQty: 0, grnQty: 0, open: 0, exp: 0, can: 0 };
    groups[key].poT += r.POTons;
    groups[key].grnT += r.GRNTons;
    groups[key].poQty += r.POQty;
    groups[key].grnQty += r.GRNQty;
    if (r.Status.toLowerCase() === 'open') groups[key].open++;
    if (r.Status.toLowerCase() === 'expired') groups[key].exp++;
    if (r.Status.toLowerCase() === 'cancelled') groups[key].can++;
  });

  let html = '';
  Object.values(groups).sort((a,b) => b.poT - a.poT).forEach(g => {
    const fillRate = g.poQty > 0 ? ((g.grnQty / g.poQty) * 100).toFixed(1) : '0.0';
    html += `<tr>
      <td><strong>${g.cat}</strong></td>
      <td>${g.city}</td>
      <td>${g.poT.toFixed(2)}</td>
      <td>${g.grnT.toFixed(2)}</td>
      <td>${fillRate}%</td>
      <td>${g.open}</td>
      <td style="color:#dc2626; font-weight:600;">${g.exp}</td>
      <td style="color:#64748b; font-weight:600;">${g.can}</td>
    </tr>`;
  });
  document.getElementById('tableBody').innerHTML = html || `<tr><td colspan="8" style="text-align:center;">No data</td></tr>`;
}

// PHASE 3: Vendor Name | Del Location | PO No | PO Date | PO Expiry | Appointment Date | Status | Sum of PO QTY | Sum of Po Value | Sum of GRN QTY | Fill rate | Month
function renderPhase3(data) {
  const cols = ['Vendor Name', 'Del Location', 'PO No', 'PO Date', 'PO Expiry', 'Appointment Date', 'Status', 'Sum of PO QTY', 'Sum of Po Value', 'Sum of GRN QTY', 'Fill rate', 'Month'];
  document.getElementById('tableHead').innerHTML = `<tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>`;
  
  // Group by PO No
  const poMap = {};
  data.forEach(r => {
    if (!poMap[r.PONo]) {
      poMap[r.PONo] = { ...r, POQty: 0, PoValue: 0, GRNQty: 0, count: 0 };
    }
    poMap[r.PONo].POQty += r.POQty;
    poMap[r.PONo].PoValue += r.PoValue;
    poMap[r.PONo].GRNQty += r.GRNQty;
    poMap[r.PONo].count++;
  });

  let html = '';
  Object.values(poMap).sort((a,b) => new Date(b.PODate) - new Date(a.PODate)).forEach(r => {
    const fillRate = r.POQty > 0 ? ((r.GRNQty / r.POQty) * 100).toFixed(1) : '0.0';
    
    // Status Badge Color
    let badgeClass = 'background:#f1f5f9; color:#475569;'; // default
    if (r.Status.toLowerCase() === 'open') badgeClass = 'background:#dcfce7; color:#166534;';
    else if (r.Status.toLowerCase() === 'expired') badgeClass = 'background:#fee2e2; color:#991b1b;';
    else if (r.Status.toLowerCase() === 'cancelled') badgeClass = 'background:#f3f4f6; color:#374151;';

    html += `<tr>
      <td>${r.VendorName}</td>
      <td>${r.DelLocation}</td>
      <td><strong>${r.PONo}</strong></td>
      <td>${r.PODate}</td>
      <td>${r.POExpiry}</td>
      <td>${r.ApptDate}</td>
      <td><span style="padding:0.25rem 0.5rem; border-radius:4px; font-size:0.75rem; font-weight:600; ${badgeClass}">${r.Status}</span></td>
      <td><strong>${r.POQty.toLocaleString()}</strong></td>
      <td>${r.PoValue.toLocaleString()}</td>
      <td>${r.GRNQty.toLocaleString()}</td>
      <td>${fillRate}%</td>
      <td>${r.Month}</td>
    </tr>`;
  });
  document.getElementById('tableBody').innerHTML = html || `<tr><td colspan="12" style="text-align:center;">No data</td></tr>`;
}

// Export CSV
document.getElementById('exportBtn').addEventListener('click', () => {
  const table = document.getElementById('dataTable');
  if (!table || table.rows.length <= 1) { alert('No data to export.'); return; }
  
  let csvContent = "";
  for (let i = 0; i < table.rows.length; i++) {
    let rowData = [];
    for (let j = 0; j < table.rows[i].cells.length; j++) {
      rowData.push(`"${table.rows[i].cells[j].innerText}"`);
    }
    csvContent += rowData.join(",") + "\n";
  }
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "PO_Export.csv";
  link.click();
});
