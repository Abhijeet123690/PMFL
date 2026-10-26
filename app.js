let rawRows = [];
let chartInstance = null;

// Set Date in header
document.getElementById('currentDate').textContent = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' }).toUpperCase();

// File Upload
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
      header: true,
      skipEmptyLines: 'greedy',
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

// Make column matching forgiving (ignores spaces, underscores, and case)
function cleanKey(str) {
  return String(str).replace(/[\s_]/g, '').toLowerCase();
}

function handleParsed(rows, fields) {
  if (!rows || rows.length === 0) {
    setStatus('The uploaded file is empty.', 'error');
    return;
  }

  // Map the actual file headers to our expected headers
  const map = {};
  fields.forEach(f => { map[cleanKey(f)] = f; });

  const required = ['Category', 'City', 'Vendor_Name', 'Delivery_Location', 'PO_No', 'PO_Date', 'PO_Expiry', 'Status', 'PO_QTY'];
  const missing = required.filter(req => !map[cleanKey(req)]);

  if (missing.length > 0) {
    setStatus('Missing columns: ' + missing.join(', ') + '. Please check your file headers.', 'error');
    return;
  }

  // Normalize data
  rawRows = rows.map(r => ({
    Category: String(r[map['category']] || '').trim(),
    City: String(r[map['city']] || '').trim(),
    Vendor_Name: String(r[map['vendorname']] || '').trim(),
    Delivery_Location: String(r[map['deliverylocation']] || '').trim(),
    PO_No: String(r[map['pono']] || '').trim(),
    PO_Date: normalizeDate(r[map['podate']]),
    PO_Expiry: normalizeDate(r[map['poexpiry']]),
    Appointment_Date: map['appointmentdate'] ? normalizeDate(r[map['appointmentdate']]) : '',
    Status: String(r[map['status']] || '').trim(),
    PO_QTY: Number(r[map['poqty']]) || 0
  })).filter(r => r.PO_No && r.Category);

  if (rawRows.length === 0) {
    setStatus('No valid rows found. Check that PO_No and Category are populated.', 'error');
    return;
  }

  setStatus(`Successfully loaded ${rawRows.length} rows.`, 'success');
  initializeDashboard();
}

function normalizeDate(val) {
  if (!val) return '';
  if (val instanceof Date) return val.toISOString().slice(0, 10);
  const s = String(val).trim();
  const d = new Date(s);
  if (!isNaN(d.getTime()) && s.match(/\d/)) return d.toISOString().slice(0, 10);
  return s;
}

function setStatus(msg, cls) {
  const el = document.getElementById('statusMsg');
  el.textContent = msg;
  el.className = 'status-msg ' + (cls || '');
}

function initializeDashboard() {
  document.getElementById('emptyState').style.display = 'none';
  document.getElementById('filterContainer').style.display = 'grid';
  document.getElementById('kpiGrid').style.display = 'grid';
  document.getElementById('mainContent').style.display = 'block';

  populateFilters();
  updateDashboard();
}

function populateFilters() {
  const cats = [...new Set(rawRows.map(r => r.Category))].sort();
  const cities = [...new Set(rawRows.map(r => r.City))].sort();
  const stats = [...new Set(rawRows.map(r => r.Status))].sort();

  document.getElementById('filterCategory').innerHTML = '<option value="">All Categories</option>' + cats.map(c => `<option value="${c}">${c}</option>`).join('');
  document.getElementById('filterCity').innerHTML = '<option value="">All Cities</option>' + cities.map(c => `<option value="${c}">${c}</option>`).join('');
  document.getElementById('filterStatus').innerHTML = '<option value="">All Statuses</option>' + stats.map(s => `<option value="${s}">${s}</option>`).join('');
}

// Event Listeners for filters
document.getElementById('filterCategory').addEventListener('change', updateDashboard);
document.getElementById('filterCity').addEventListener('change', updateDashboard);
document.getElementById('filterStatus').addEventListener('change', updateDashboard);
document.getElementById('searchInput').addEventListener('input', updateDashboard);
document.getElementById('resetBtn').addEventListener('click', () => {
  document.getElementById('filterCategory').value = '';
  document.getElementById('filterCity').value = '';
  document.getElementById('filterStatus').value = '';
  document.getElementById('searchInput').value = '';
  updateDashboard();
});

function updateDashboard() {
  const catFilter = document.getElementById('filterCategory').value;
  const cityFilter = document.getElementById('filterCity').value;
  const statusFilter = document.getElementById('filterStatus').value;
  const searchStr = document.getElementById('searchInput').value.toLowerCase();

  // Apply Filters
  const filtered = rawRows.filter(r => {
    if (catFilter && r.Category !== catFilter) return false;
    if (cityFilter && r.City !== cityFilter) return false;
    if (statusFilter && r.Status !== statusFilter) return false;
    if (searchStr && !r.PO_No.toLowerCase().includes(searchStr) && !r.Vendor_Name.toLowerCase().includes(searchStr)) return false;
    return true;
  });

  document.getElementById('rowCount').textContent = `${filtered.length} filtered rows`;

  // Update KPIs
  const uniquePOs = new Set(filtered.map(r => r.PO_No)).size;
  const totalQty = filtered.reduce((sum, r) => sum + r.PO_QTY, 0);
  const openCount = filtered.filter(r => r.Status.toLowerCase() === 'open').length;
  const expiredCount = filtered.filter(r => r.Status.toLowerCase() === 'expired').length;
  
  // Expiring within 30 days
  const expiring = filtered.filter(r => {
    if (!r.PO_Expiry) return false;
    const days = Math.ceil((new Date(r.PO_Expiry) - new Date()) / (1000 * 60 * 60 * 24));
    return days >= 0 && days <= 30 && r.Status !== 'Delivered' && r.Status !== 'Cancelled';
  }).length;

  document.getElementById('kpiTotalPOs').textContent = uniquePOs.toLocaleString();
  document.getElementById('kpiTotalQty').textContent = totalQty.toLocaleString();
  document.getElementById('kpiOpen').textContent = openCount.toLocaleString();
  document.getElementById('kpiExpired').textContent = expiredCount.toLocaleString();
  document.getElementById('kpiExpiring').textContent = expiring.toLocaleString();

  // Update Table Title
  let title = "PO Tracking — ";
  if (catFilter && cityFilter) title += `${catFilter} in ${cityFilter}`;
  else if (catFilter) title += catFilter;
  else if (cityFilter) title += cityFilter;
  else title += "Overview";
  document.getElementById('viewTitle').textContent = title;

  renderTable(filtered);
}

function renderTable(data) {
  const thead = document.getElementById('tableHead');
  const tbody = document.getElementById('tableBody');

  // Dynamically change columns based on drill-down level
  const catFilter = document.getElementById('filterCategory').value;
  const cityFilter = document.getElementById('filterCity').value;

  let cols = [];
  if (!catFilter && !cityFilter) {
    // Phase 1: Overall
    cols = ['Category', 'Total POs', 'Total PO QTY', 'Open', 'Expired'];
  } else if (catFilter && !cityFilter) {
    // Phase 2: Category + City
    cols = ['City', 'Total POs', 'Total PO QTY', 'Open', 'Expired'];
  } else {
    // Phase 3: Detailed Lines
    cols = ['Vendor Name', 'Delivery Location', 'PO No', 'PO Date', 'PO Expiry', 'Appointment Date', 'Status', 'PO QTY'];
  }

  thead.innerHTML = `<tr>${cols.map(c => `<th>${c.toUpperCase()}</th>`).join('')}</tr>`;

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${cols.length}" style="text-align:center; padding: 2rem;">No data matches your filters.</td></tr>`;
    return;
  }

  let html = '';

  // Aggregate logic depending on view
  if (!catFilter && !cityFilter) {
    const groups = {};
    data.forEach(r => {
      if (!groups[r.Category]) groups[r.Category] = [];
      groups[r.Category].push(r);
    });
    Object.keys(groups).sort().forEach(cat => {
      const rows = groups[cat];
      const pos = new Set(rows.map(r => r.PO_No)).size;
      const qty = rows.reduce((s, r) => s + r.PO_QTY, 0);
      const open = rows.filter(r => r.Status.toLowerCase() === 'open').length;
      const exp = rows.filter(r => r.Status.toLowerCase() === 'expired').length;
      html += `<tr>
        <td><strong>${cat}</strong></td>
        <td>${pos}</td>
        <td>${qty.toLocaleString()}</td>
        <td>${open}</td>
        <td>${exp}</td>
      </tr>`;
    });
  } else if (catFilter && !cityFilter) {
    const groups = {};
    data.forEach(r => {
      if (!groups[r.City]) groups[r.City] = [];
      groups[r.City].push(r);
    });
    Object.keys(groups).sort().forEach(city => {
      const rows = groups[city];
      const pos = new Set(rows.map(r => r.PO_No)).size;
      const qty = rows.reduce((s, r) => s + r.PO_QTY, 0);
      const open = rows.filter(r => r.Status.toLowerCase() === 'open').length;
      const exp = rows.filter(r => r.Status.toLowerCase() === 'expired').length;
      html += `<tr>
        <td><strong>${city}</strong></td>
        <td>${pos}</td>
        <td>${qty.toLocaleString()}</td>
        <td>${open}</td>
        <td>${exp}</td>
      </tr>`;
    });
  } else {
    // Detailed View (Aggregated by PO)
    const poMap = new Map();
    data.forEach(r => {
      if (!poMap.has(r.PO_No)) poMap.set(r.PO_No, { ...r, PO_QTY: 0 });
      poMap.get(r.PO_No).PO_QTY += r.PO_QTY;
    });

    Array.from(poMap.values()).forEach(r => {
      let badgeClass = 'badge-delivered';
      if (r.Status.toLowerCase() === 'open') badgeClass = 'badge-open';
      else if (r.Status.toLowerCase() === 'expired') badgeClass = 'badge-expired';
      else if (r.Status.toLowerCase() === 'in transit') badgeClass = 'badge-transit';

      html += `<tr>
        <td>${r.Vendor_Name}</td>
        <td>${r.Delivery_Location}</td>
        <td><strong>${r.PO_No}</strong></td>
        <td>${r.PO_Date || '—'}</td>
        <td>${r.PO_Expiry || '—'}</td>
        <td>${r.Appointment_Date || '—'}</td>
        <td><span class="badge ${badgeClass}">${r.Status}</span></td>
        <td><strong>${r.PO_QTY.toLocaleString()}</strong></td>
      </tr>`;
    });
  }

  tbody.innerHTML = html;
}

// Export Function
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
