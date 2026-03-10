
let deferredPrompt = null;
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js').catch(err => console.log(err));
  });
}
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredPrompt = e;
  const b = document.getElementById('installBtn');
  if (b) b.style.display = 'inline-flex';
});
document.addEventListener('click', async e => {
  if (e.target && e.target.id === 'installBtn' && deferredPrompt) {
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    deferredPrompt = null;
    e.target.style.display = 'none';
  }
});

let inventoryData = [];
let stocktake = {};
let settings = { lowStockThreshold: 10, sessionName: '' };
let db;
const PAGE_META = {
  upload: { title: 'Upload Inventory', subtitle: 'Import Excel data, review it, and prepare your offline workspace.', actionLabel: 'Load sample', actionId: 'loadSampleBtn' },
  dashboard: { title: 'Dashboard', subtitle: 'See stock health, count progress, and the most urgent lines at a glance.', actionLabel: 'Open alerts', actionId: 'goToAlertsBtn' },
  alerts: { title: 'Alerts', subtitle: 'Focus on stockouts, low stock items, and lines that need review.', actionLabel: 'Export alerts', actionId: 'exportAlertsBtn' },
  items: { title: 'Manage Items', subtitle: 'Add, update, and maintain inventory records without leaving the page.', actionLabel: 'Save item', actionId: 'addItemBtn' },
  stocktake: { title: 'Stocktake', subtitle: 'Count items, capture variances, and keep the session updated in real time.', actionLabel: 'Export variances', actionId: 'exportVarianceBtn' },
  settings: { title: 'Settings', subtitle: 'Install the app, export files, and manage local workspace behavior.', actionLabel: 'Export cycle count sheet', actionId: 'exportCycleCountBtn' }
};
const $ = id => document.getElementById(id);
const $$ = selector => Array.from(document.querySelectorAll(selector));

function toNumberOrNull(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'string') {
    const t = v.trim();
    if (!t || t.toUpperCase() === '#N/A') return null;
    const n = Number(t.replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function cleanText(v) {
  if (v === null || v === undefined) return '';
  const t = String(v).trim();
  return t.toUpperCase() === '#N/A' ? '' : t;
}
function formatNumber(n) {
  return Number(n || 0).toLocaleString();
}
function getStatus(item) {
  const qty = Number(item.QtyOnHand || 0);
  const min = item.MinLevel;
  if (qty <= 0) return 'STOCKOUT';
  if (min !== null && min !== undefined && min !== '' && qty < Number(min)) return 'LOW STOCK';
  if (qty < Number(settings.lowStockThreshold || 0)) return 'REVIEW';
  return 'OK';
}
function statusBadge(s) {
  if (s === 'STOCKOUT') return '<span class="pill pill-danger">Stockout</span>';
  if (s === 'LOW STOCK') return '<span class="pill pill-warning">Low stock</span>';
  if (s === 'REVIEW') return '<span class="pill pill-warning">Review</span>';
  return '<span class="pill pill-ok">OK</span>';
}
function rowClass(s) {
  if (s === 'STOCKOUT') return 'stockout';
  if (s === 'LOW STOCK' || s === 'REVIEW') return 'low-stock';
  return 'ok-row';
}
function normalizeItem(raw) {
  const stock = raw.StockCode ?? raw['Stock Code'] ?? raw['SAP Material'];
  const desc = raw.Description ?? raw['Material Description'] ?? raw['Item Description'];
  if (!stock || !desc) return null;
  return {
    StockCode: cleanText(stock),
    Description: cleanText(desc),
    AlternateKey: cleanText(raw.AlternateKey ?? raw['old stkcode']),
    QtyOnHand: toNumberOrNull(raw.QtyOnHand ?? raw['Stock on hand '] ?? raw['Stock on hand'] ?? raw['Qty']) ?? 0,
    DefaultBin: cleanText(raw.DefaultBin ?? raw['Bin location'] ?? raw['Default Bin'] ?? raw['Bin']),
    MinLevel: toNumberOrNull(raw.MinLevel ?? raw['minimum QTY'] ?? raw['Min']),
    MaxLevel: toNumberOrNull(raw.MaxLevel ?? raw['maximum QTY'] ?? raw['Max']),
    UnitCost: toNumberOrNull(raw.UnitCost ?? raw['Unit cost '] ?? raw['Unit cost']),
    TotalValue: toNumberOrNull(raw.TotalValue ?? raw['Total value']),
    Department: cleanText(raw.Department ?? raw['Department ']),
    Machine: cleanText(raw.Machine ?? raw['Machine ']),
    Currency: cleanText(raw.Currency)
  };
}
function downloadExcel(rows, sheetName, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}
function showToast(title, text = '', type = 'info') {
  const root = $('toastRoot');
  if (!root) return;
  const iconMap = { success: '✓', error: '⚠', warning: '!', info: 'ℹ' };
  const toast = document.createElement('div');
  toast.className = `toast toast--${type}`;
  toast.innerHTML = `<div class="toast__icon">${iconMap[type] || 'ℹ'}</div><div><div class="toast__title">${escapeHtml(title)}</div>${text ? `<div class="toast__text">${escapeHtml(text)}</div>` : ''}</div>`;
  root.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(-6px)';
    setTimeout(() => toast.remove(), 220);
  }, 3600);
}
function closeModal() {
  const root = $('systemModalRoot');
  root.classList.remove('open');
  root.innerHTML = '';
}
function openModal({ title, message, type = 'info', confirmText = 'Continue', cancelText = 'Cancel', hideCancel = false, onConfirm }) {
  const root = $('systemModalRoot');
  const iconClass = type === 'danger' ? 'danger' : 'info';
  root.classList.add('open');
  root.innerHTML = `
    <div class="modal-backdrop" data-close-modal></div>
    <div class="modal" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
      <div class="modal-icon ${iconClass}">${type === 'danger' ? '⚠' : 'ℹ'}</div>
      <div class="eyebrow">System message</div>
      <h3 id="modalTitle">${escapeHtml(title)}</h3>
      <p>${escapeHtml(message)}</p>
      <div class="modal-actions">
        ${hideCancel ? '' : `<button class="btn btn-secondary" type="button" id="modalCancelBtn">${escapeHtml(cancelText)}</button>`}
        <button class="btn ${type === 'danger' ? 'btn-primary' : 'btn-primary'}" type="button" id="modalConfirmBtn">${escapeHtml(confirmText)}</button>
      </div>
    </div>`;
  const confirmBtn = $('modalConfirmBtn');
  const cancelBtn = $('modalCancelBtn');
  confirmBtn?.focus();
  confirmBtn?.addEventListener('click', () => {
    closeModal();
    if (typeof onConfirm === 'function') onConfirm();
  });
  cancelBtn?.addEventListener('click', closeModal);
  root.querySelector('[data-close-modal]')?.addEventListener('click', closeModal);
}

const request = indexedDB.open('InventoryPWA_DB_MOBILE', 1);
request.onerror = e => console.log('IndexedDB error:', e.target.errorCode);
request.onupgradeneeded = e => {
  db = e.target.result;
  if (!db.objectStoreNames.contains('inventory')) db.createObjectStore('inventory', { keyPath: 'StockCode' });
  if (!db.objectStoreNames.contains('stocktake')) db.createObjectStore('stocktake', { keyPath: 'StockCode' });
  if (!db.objectStoreNames.contains('settings')) db.createObjectStore('settings', { keyPath: 'key' });
};
request.onsuccess = e => {
  db = e.target.result;
  loadAll();
};
function tx(name, mode = 'readonly') {
  return db.transaction([name], mode).objectStore(name);
}
function saveInventory() {
  const s = tx('inventory', 'readwrite');
  s.clear();
  inventoryData.forEach(it => s.put(it));
}
function saveStocktake() {
  const s = tx('stocktake', 'readwrite');
  s.clear();
  Object.entries(stocktake).forEach(([StockCode, obj]) => s.put({ StockCode, ...obj }));
}
function saveSettings() {
  const s = tx('settings', 'readwrite');
  s.put({ key: 'lowStockThreshold', value: settings.lowStockThreshold });
  s.put({ key: 'sessionName', value: settings.sessionName });
}
function loadAll() {
  tx('inventory').getAll().onsuccess = e => {
    inventoryData = e.target.result || [];
    tx('stocktake').getAll().onsuccess = e2 => {
      stocktake = {};
      (e2.target.result || []).forEach(r => stocktake[r.StockCode] = { CountedQty: r.CountedQty ?? null, UpdatedAt: r.UpdatedAt ?? null });
      tx('settings').getAll().onsuccess = e3 => {
        (e3.target.result || []).forEach(r => settings[r.key] = r.value);
        $('lowStockThreshold').value = settings.lowStockThreshold ?? 10;
        $('sessionName').value = settings.sessionName ?? '';
        renderAll();
        updateTopBarForTab('upload');
        syncSessionChips();
      };
    };
  };
}

function setActiveTab(tab) {
  $$('.nav-link, .mobile-nav__link').forEach(btn => btn.classList.toggle('active', btn.dataset.tab === tab));
  $$('.tab-section').forEach(sec => sec.hidden = sec.id !== tab);
  updateTopBarForTab(tab);
  renderAll();
}
function updateTopBarForTab(tab) {
  const meta = PAGE_META[tab] || PAGE_META.upload;
  $('pageTitle').textContent = meta.title;
  $('pageSubtitle').textContent = meta.subtitle;
  $('topQuickAction').textContent = meta.actionLabel;
  $('topQuickAction').onclick = () => document.getElementById(meta.actionId)?.click();
}
function syncSessionChips() {
  const sessionText = settings.sessionName ? `Session: ${settings.sessionName}` : 'No active session';
  $('sessionChip').textContent = sessionText;
  $('dashboardSessionName').textContent = settings.sessionName || 'No session saved';
  $('dashboardSessionMeta').textContent = settings.sessionName
    ? 'Exports from the stocktake section will include this session name.'
    : 'Save a stocktake session name to label exports and keep your work organized.';
}
function updateConnectionState() {
  const online = navigator.onLine;
  const dot = $('deviceStatusDot');
  const label = $('deviceStatusText');
  dot.classList.toggle('is-online', online);
  dot.classList.toggle('is-offline', !online);
  label.textContent = online ? 'Online and ready to sync files' : 'Offline mode active';
}

$$('.nav-link, .mobile-nav__link').forEach(btn => btn.addEventListener('click', () => setActiveTab(btn.dataset.tab)));
$('goToAlertsBtn').addEventListener('click', () => setActiveTab('alerts'));
$('goToStocktakeBtn').addEventListener('click', () => setActiveTab('stocktake'));
window.addEventListener('online', updateConnectionState);
window.addEventListener('offline', updateConnectionState);
updateConnectionState();

$('excelUpload').addEventListener('change', event => {
  const file = event.target.files[0];
  if (!file) return;
  showToast('Import started', `Reading ${file.name}`, 'info');
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const preferredSheet = workbook.SheetNames.find(n => n.trim().toUpperCase() === 'MR01') || workbook.SheetNames[0];
      const sheet = workbook.Sheets[preferredSheet];
      const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      const normalized = raw.map(normalizeItem).filter(Boolean);
      const existing = { ...stocktake };
      inventoryData = normalized;
      stocktake = {};
      inventoryData.forEach(it => { if (existing[it.StockCode]) stocktake[it.StockCode] = existing[it.StockCode]; });
      saveInventory();
      saveStocktake();
      renderAll();
      showToast('Import complete', `Loaded ${formatNumber(normalized.length)} items from ${preferredSheet}.`, 'success');
    } catch (err) {
      console.error(err);
      openModal({ title: 'Import failed', message: 'The file could not be imported. Use the current SAP export structure and try again.', type: 'danger', confirmText: 'Close', hideCancel: true });
    }
  };
  reader.readAsArrayBuffer(file);
});
$('loadSampleBtn').addEventListener('click', async () => {
  const res = await fetch('./sample-data.json');
  const raw = await res.json();
  inventoryData = raw.map(normalizeItem).filter(Boolean);
  stocktake = {};
  saveInventory();
  saveStocktake();
  renderAll();
  showToast('Sample data loaded', 'You can now review and test the redesigned interface.', 'success');
});
$('searchInput').addEventListener('input', renderPreviewTable);
$('stocktakeFilter').addEventListener('input', renderStocktakeTable);
$('alertSearch').addEventListener('input', renderAlertsTable);
$('alertFilter').addEventListener('change', () => {
  syncAlertChips();
  renderAlertsTable();
});
$$('[data-alert-filter]').forEach(btn => btn.addEventListener('click', () => {
  $('alertFilter').value = btn.dataset.alertFilter;
  syncAlertChips();
  renderAlertsTable();
}));
$('lowStockThreshold').addEventListener('input', () => {
  settings.lowStockThreshold = Number($('lowStockThreshold').value || 0);
  saveSettings();
  renderAll();
});
$('clearAllBtn').addEventListener('click', () => {
  openModal({
    title: 'Clear local workspace?',
    message: 'This will remove all inventory, stocktake counts, and the saved session from this device.',
    type: 'danger',
    confirmText: 'Clear data',
    onConfirm: () => {
      inventoryData = [];
      stocktake = {};
      settings.sessionName = '';
      saveInventory();
      saveStocktake();
      saveSettings();
      $('sessionName').value = '';
      renderAll();
      syncSessionChips();
      showToast('Workspace cleared', 'All local data has been removed from this device.', 'warning');
    }
  });
});
$('exportInventoryBtn').addEventListener('click', () => {
  if (inventoryData.length === 0) return showToast('Nothing to export', 'Load inventory data first.', 'warning');
  downloadExcel(inventoryData, 'Inventory', 'InventoryData_CurrentFormat.xlsx');
  showToast('Inventory exported', 'The current inventory file has been downloaded.', 'success');
});
$('exportAlertsBtn').addEventListener('click', exportAlertsExcel);
$('addItemBtn').addEventListener('click', () => {
  const newItem = normalizeItem({
    StockCode: $('newStockCode').value,
    Description: $('newDescription').value,
    'old stkcode': $('newAlternateKey').value,
    'Stock on hand': $('newQtyOnHand').value,
    'Bin location': $('newDefaultBin').value,
    'minimum QTY': $('newMinLevel').value,
    'maximum QTY': $('newMaxLevel').value
  });
  if (!newItem) {
    openModal({ title: 'Missing required details', message: 'Material and Description are required before saving an item.', type: 'info', confirmText: 'Got it', hideCancel: true });
    return;
  }
  const idx = inventoryData.findIndex(x => x.StockCode === newItem.StockCode);
  if (idx >= 0) inventoryData[idx] = { ...inventoryData[idx], ...newItem };
  else inventoryData.push(newItem);
  saveInventory();
  renderAll();
  ['newStockCode', 'newDescription', 'newAlternateKey', 'newDefaultBin', 'newMinLevel', 'newMaxLevel'].forEach(id => $(id).value = '');
  $('newQtyOnHand').value = '0';
  showToast(idx >= 0 ? 'Item updated' : 'Item added', `${newItem.StockCode} has been saved successfully.`, 'success');
});
$('saveSessionNameBtn').addEventListener('click', () => {
  settings.sessionName = $('sessionName').value.trim();
  saveSettings();
  syncSessionChips();
  showToast('Session saved', settings.sessionName || 'Session name cleared.', 'success');
});
$('markAllUncountedBtn').addEventListener('click', () => {
  openModal({
    title: 'Clear all counted quantities?',
    message: 'Every counted quantity will be reset to blank, but inventory records will remain.',
    type: 'danger',
    confirmText: 'Reset counts',
    onConfirm: () => {
      stocktake = {};
      saveStocktake();
      renderAll();
      showToast('Counts cleared', 'All counted quantities have been reset.', 'warning');
    }
  });
});
$('exportVarianceBtn').addEventListener('click', exportVarianceExcel);
$('exportFullCountBtn').addEventListener('click', exportFullCountExcel);
$('exportCycleCountBtn').addEventListener('click', exportCycleCountSheet);
$('exportStocktakeBtn').addEventListener('click', exportVarianceExcel);

function matchesQuery(it, q) {
  if (!q) return true;
  return ((it.StockCode || '').toLowerCase().includes(q)
    || (it.Description || '').toLowerCase().includes(q)
    || (it.AlternateKey || '').toLowerCase().includes(q)
    || (it.Department || '').toLowerCase().includes(q)
    || (it.Machine || '').toLowerCase().includes(q));
}
function alertItems() {
  return inventoryData.filter(it => getStatus(it) !== 'OK');
}
function renderAll() {
  renderPreviewTable();
  renderItemsTable();
  renderStocktakeTable();
  renderAlertsTable();
  renderDashboardAlertsTable();
  updateDashboard();
  syncSessionChips();
  syncAlertChips();
}
function renderPreviewTable() {
  const tbody = document.querySelector('#previewTable tbody');
  tbody.innerHTML = '';
  const q = $('searchInput').value.toLowerCase().trim();
  const rows = inventoryData.filter(it => matchesQuery(it, q)).slice(0, 250);
  if (rows.length === 0) return emptyTable(tbody, 6, 'No inventory rows to preview yet.');
  rows.forEach(it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td class="right mono">${it.MinLevel ?? ''}</td><td>${statusBadge(s)}</td>`;
    tbody.appendChild(tr);
  });
}
function renderDashboardAlertsTable() {
  const tbody = document.querySelector('#dashboardAlertsTable tbody');
  tbody.innerHTML = '';
  const order = { 'STOCKOUT': 0, 'LOW STOCK': 1, 'REVIEW': 2 };
  const rows = alertItems().slice().sort((a, b) => order[getStatus(a)] - order[getStatus(b)]).slice(0, 10);
  if (rows.length === 0) return emptyTable(tbody, 6, 'No critical alerts at the moment.');
  rows.forEach(it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right mono">${it.MinLevel ?? ''}</td><td class="hide-mobile">${escapeHtml(it.Department || '')}</td><td>${statusBadge(s)}</td>`;
    tbody.appendChild(tr);
  });
}
function renderAlertsTable() {
  const tbody = document.querySelector('#alertsTable tbody');
  tbody.innerHTML = '';
  const filter = $('alertFilter').value;
  const q = $('alertSearch').value.toLowerCase().trim();
  const order = { 'STOCKOUT': 0, 'LOW STOCK': 1, 'REVIEW': 2 };
  const rows = alertItems()
    .filter(it => matchesQuery(it, q))
    .filter(it => {
      const s = getStatus(it);
      if (filter === 'stockout') return s === 'STOCKOUT';
      if (filter === 'low') return s === 'LOW STOCK';
      if (filter === 'review') return s === 'REVIEW';
      return true;
    })
    .sort((a, b) => order[getStatus(a)] - order[getStatus(b)]);
  if (rows.length === 0) return emptyTable(tbody, 7, 'No alerts match the current filter.');
  rows.forEach(it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right mono">${it.MinLevel ?? ''}</td><td class="hide-mobile">${escapeHtml(it.Department || '')}</td><td class="hide-mobile">${escapeHtml(it.Machine || '')}</td><td>${statusBadge(s)}</td>`;
    tbody.appendChild(tr);
  });
}
function renderItemsTable() {
  const tbody = document.querySelector('#itemsTable tbody');
  tbody.innerHTML = '';
  $('itemsCountLabel').textContent = `${formatNumber(inventoryData.length)} items`;
  const rows = [...inventoryData].sort((a, b) => String(a.StockCode).localeCompare(String(b.StockCode)));
  if (rows.length === 0) return emptyTable(tbody, 6, 'No items have been loaded yet.');
  rows.forEach(it => {
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td>${statusBadge(s)}</td><td class="center"><button data-del="${escapeAttr(it.StockCode)}" class="btn btn-secondary" style="min-height:40px;padding:0 12px;">Delete</button></td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('button[data-del]').forEach(btn => btn.addEventListener('click', () => {
    const code = btn.getAttribute('data-del');
    openModal({
      title: `Delete ${code}?`,
      message: 'This item and its related stocktake count will be removed from this device.',
      type: 'danger',
      confirmText: 'Delete item',
      onConfirm: () => {
        inventoryData = inventoryData.filter(x => x.StockCode !== code);
        delete stocktake[code];
        saveInventory();
        saveStocktake();
        renderAll();
        showToast('Item deleted', `${code} has been removed.`, 'success');
      }
    });
  }));
}
function renderStocktakeTable() {
  const tbody = document.querySelector('#stocktakeTable tbody');
  tbody.innerHTML = '';
  const q = $('stocktakeFilter').value.toLowerCase().trim();
  const rows = [...inventoryData]
    .filter(it => matchesQuery(it, q))
    .sort((a, b) => String(a.StockCode).localeCompare(String(b.StockCode)));
  if (rows.length === 0) return emptyTable(tbody, 7, 'No items available for stocktake.');
  rows.forEach(it => {
    const counted = stocktake[it.StockCode]?.CountedQty ?? null;
    const variance = counted === null ? '' : (counted - (it.QtyOnHand || 0));
    const s = getStatus(it);
    const tr = document.createElement('tr');
    tr.className = rowClass(s);
    tr.innerHTML = `<td class="mono">${escapeHtml(it.StockCode)}</td><td>${escapeHtml(it.Description)}</td><td class="right mono">${it.QtyOnHand ?? 0}</td><td class="right"><input inputmode="numeric" style="max-width:120px;margin-left:auto;" type="number" data-count="${escapeAttr(it.StockCode)}" value="${counted ?? ''}" /></td><td class="right mono" data-var="${escapeAttr(it.StockCode)}">${variance}</td><td class="hide-mobile">${escapeHtml(it.DefaultBin || '')}</td><td>${statusBadge(s)}</td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll('input[data-count]').forEach(inp => inp.addEventListener('input', () => {
    const code = inp.getAttribute('data-count');
    const countedQty = toNumberOrNull(inp.value);
    if (countedQty === null) delete stocktake[code];
    else stocktake[code] = { CountedQty: countedQty, UpdatedAt: new Date().toISOString() };
    saveStocktake();
    const it = inventoryData.find(x => x.StockCode === code);
    const varCell = tbody.querySelector(`[data-var="${cssEscape(code)}"]`);
    if (varCell) varCell.textContent = countedQty === null ? '' : (countedQty - (it?.QtyOnHand || 0));
    inp.closest('tr')?.classList.add('row-saved');
    setTimeout(() => inp.closest('tr')?.classList.remove('row-saved'), 600);
    updateDashboard();
  }));
}
function updateDashboard() {
  const totalItems = inventoryData.length;
  const flagged = alertItems().length;
  const counted = Object.keys(stocktake).length;
  const varianceLines = Object.keys(stocktake).filter(code => {
    const it = inventoryData.find(x => x.StockCode === code);
    const countedQty = stocktake[code]?.CountedQty;
    return countedQty !== null && countedQty !== undefined && (countedQty - (it?.QtyOnHand || 0)) !== 0;
  }).length;
  $('totalItems').innerText = formatNumber(totalItems);
  $('totalQty').innerText = formatNumber(inventoryData.reduce((s, it) => s + (it.QtyOnHand || 0), 0));
  $('stockoutCount').innerText = formatNumber(inventoryData.filter(it => getStatus(it) === 'STOCKOUT').length);
  $('lowStockCount').innerText = formatNumber(inventoryData.filter(it => getStatus(it) === 'LOW STOCK').length);
  $('inventoryValue').innerText = formatNumber(inventoryData.reduce((s, it) => s + (it.TotalValue || ((it.UnitCost || 0) * (it.QtyOnHand || 0))), 0));
  $('stockoutValue').innerText = formatNumber(inventoryData.filter(it => getStatus(it) === 'STOCKOUT').reduce((s, it) => s + (it.UnitCost || 0), 0));
  $('countedLines').innerText = formatNumber(counted);
  $('varianceLines').innerText = formatNumber(varianceLines);
  $('heroImportedItems').innerText = formatNumber(totalItems);
  $('heroFlaggedItems').innerText = formatNumber(flagged);
  $('countedProgressText').textContent = `${formatNumber(counted)} of ${formatNumber(totalItems)} lines counted`;
  $('flaggedProgressText').textContent = `${formatNumber(flagged)} critical items need attention`;
  $('countedProgressBar').style.width = totalItems ? `${(counted / totalItems) * 100}%` : '0%';
  $('flaggedProgressBar').style.width = totalItems ? `${(flagged / totalItems) * 100}%` : '0%';
}
function syncAlertChips() {
  const value = $('alertFilter').value;
  $$('[data-alert-filter]').forEach(btn => btn.classList.toggle('active', btn.dataset.alertFilter === value));
}
function emptyTable(tbody, colSpan, message) {
  const tr = document.createElement('tr');
  tr.innerHTML = `<td colspan="${colSpan}" class="helper-empty">${escapeHtml(message)}</td>`;
  tbody.appendChild(tr);
}
function exportVarianceExcel() {
  if (inventoryData.length === 0) return showToast('No inventory loaded', 'Import or create inventory items first.', 'warning');
  const rows = [];
  Object.entries(stocktake).forEach(([code, obj]) => {
    const it = inventoryData.find(x => x.StockCode === code);
    if (!it || obj.CountedQty === null || obj.CountedQty === undefined) return;
    const variance = obj.CountedQty - (it.QtyOnHand || 0);
    if (variance === 0) return;
    rows.push({
      Session: settings.sessionName || '',
      SAP_Material: it.StockCode,
      Material_Description: it.Description,
      Bin_location: it.DefaultBin || '',
      Department: it.Department || '',
      Machine: it.Machine || '',
      SystemQty: it.QtyOnHand || 0,
      CountedQty: obj.CountedQty,
      Variance: variance,
      Status: getStatus(it),
      UpdatedAt: obj.UpdatedAt || ''
    });
  });
  if (rows.length === 0) return showToast('No variances found', 'Update counted quantities to create a variance report.', 'warning');
  downloadExcel(rows, 'Variances', 'Stocktake_Variances.xlsx');
  showToast('Variance export ready', 'The variance report has been downloaded.', 'success');
}
function exportFullCountExcel() {
  if (inventoryData.length === 0) return showToast('No inventory loaded', 'Import or create inventory items first.', 'warning');
  const rows = inventoryData.map(it => {
    const counted = stocktake[it.StockCode]?.CountedQty ?? null;
    return {
      Session: settings.sessionName || '',
      SAP_Material: it.StockCode,
      Material_Description: it.Description,
      Bin_location: it.DefaultBin || '',
      Department: it.Department || '',
      Machine: it.Machine || '',
      Stock_on_hand: it.QtyOnHand || 0,
      minimum_QTY: it.MinLevel ?? '',
      maximum_QTY: it.MaxLevel ?? '',
      Status: getStatus(it),
      CountedQty: counted,
      Variance: counted === null ? null : (counted - (it.QtyOnHand || 0))
    };
  });
  downloadExcel(rows, 'FullCount', 'Stocktake_FullCount.xlsx');
  showToast('Full count export ready', 'The stocktake full count file has been downloaded.', 'success');
}
function exportCycleCountSheet() {
  if (inventoryData.length === 0) return showToast('No inventory loaded', 'Import or create inventory items first.', 'warning');
  const rows = inventoryData.map(it => ({
    SAP_Material: it.StockCode,
    Material_Description: it.Description,
    Bin_location: it.DefaultBin || '',
    Department: it.Department || '',
    Machine: it.Machine || '',
    Stock_on_hand: it.QtyOnHand || 0,
    CountedQty: '',
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    Status: getStatus(it)
  }));
  downloadExcel(rows, 'CycleCount', 'CycleCountSheet.xlsx');
  showToast('Cycle count sheet ready', 'The blank cycle count sheet has been downloaded.', 'success');
}
function exportAlertsExcel() {
  const rows = alertItems().map(it => ({
    SAP_Material: it.StockCode,
    Material_Description: it.Description,
    Bin_location: it.DefaultBin || '',
    Department: it.Department || '',
    Machine: it.Machine || '',
    Stock_on_hand: it.QtyOnHand || 0,
    minimum_QTY: it.MinLevel ?? '',
    maximum_QTY: it.MaxLevel ?? '',
    Status: getStatus(it)
  }));
  if (rows.length === 0) return showToast('No alerts found', 'There are no flagged stock lines to export.', 'warning');
  downloadExcel(rows, 'StockAlerts', 'Stock_Alerts_Report.xlsx');
  showToast('Alerts exported', 'The stock alerts report has been downloaded.', 'success');
}
function escapeHtml(str) {
  return String(str).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
}
function escapeAttr(str) { return escapeHtml(str); }
function cssEscape(str) { return String(str).replace(/"/g, '\\"'); }
