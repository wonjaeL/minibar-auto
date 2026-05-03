// ---------- Tabs ----------
document.querySelectorAll('.tab').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((b) => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
  });
});

// ---------- Excel tab ----------
const excelState = { path: null };
const $excelPath = document.getElementById('excel-path');
const $excelDate = document.getElementById('excel-date');
const $excelRun = document.getElementById('excel-run');
const $excelCopy = document.getElementById('excel-copy');
const $excelOut = document.getElementById('excel-out');
const $excelMeta = document.getElementById('excel-meta');

function refreshExcelRun() {
  $excelRun.disabled = !(excelState.path && $excelDate.value);
}

document.getElementById('excel-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('excel');
  if (!p) return;
  excelState.path = p;
  $excelPath.textContent = p;
  $excelPath.classList.remove('muted');
  refreshExcelRun();
});

$excelDate.addEventListener('change', refreshExcelRun);

$excelRun.addEventListener('click', async () => {
  const day = parseInt($excelDate.value.split('-')[2], 10);
  $excelMeta.textContent = '추출 중…';
  const r = await window.api.getRooms(excelState.path, day);
  if (r.error) {
    $excelOut.value = '';
    $excelMeta.textContent = `오류: ${r.error}` + (r.sheets ? ` | 시트: ${r.sheets.join(', ')}` : '');
    $excelCopy.disabled = true;
    return;
  }
  $excelOut.value = r.csv;
  $excelMeta.textContent = `${r.rooms.length}개 room (sheet: ${r.sheets.join(', ')})`;
  $excelCopy.disabled = r.rooms.length === 0;
});

$excelCopy.addEventListener('click', async () => {
  await navigator.clipboard.writeText($excelOut.value);
  $excelMeta.textContent = '클립보드에 복사됨';
});

// ---------- XML tab ----------
const xmlState = {
  path: null,
  keys: [],         // all available column keys
  rows: [],         // raw flattened records
  visibleKeys: [],  // user-selected columns
  filters: {},      // {key: text}
  sortBy: null,
  sortDir: 'asc',
};

const $xmlPath = document.getElementById('xml-path');
const $xmlCount = document.getElementById('xml-count');
const $xmlColumns = document.getElementById('xml-columns');
const $xmlTable = document.getElementById('xml-table');

const DEFAULT_COLUMNS = [
  'BUSINESS_DATE', 'BUSINESS_TIME', 'ROOM', 'ARTICLE_DESCRIPTION', 'QUANTITY',
  'CASHIER_DEBIT', 'CASHIER_CREDIT', 'GUEST_FULL_NAME', 'TRX_NO', 'TRX_DESC',
  'MARKET_CODE', 'ROOM_CLASS', 'CASHIER_NAME',
];

document.getElementById('xml-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('xml');
  if (!p) return;
  xmlState.path = p;
  $xmlPath.textContent = p;
  $xmlPath.classList.remove('muted');
  $xmlCount.textContent = '파싱 중…';
  const r = await window.api.parseXml(p);
  if (r.error) {
    $xmlCount.textContent = `오류: ${r.error}`;
    return;
  }
  xmlState.keys = r.keys;
  xmlState.rows = r.rows;
  xmlState.visibleKeys = DEFAULT_COLUMNS.filter((k) => r.keys.includes(k));
  if (xmlState.visibleKeys.length === 0) xmlState.visibleKeys = r.keys.slice(0, 10);
  xmlState.filters = {};
  xmlState.sortBy = null;
  $xmlCount.textContent = `${r.count}개 레코드`;
  renderColumnPicker();
  renderTable();
});

function renderColumnPicker() {
  // Replace single-select with a checkbox group inline
  const wrap = document.createElement('div');
  wrap.className = 'col-picker';
  for (const k of xmlState.keys) {
    const id = 'col-' + k;
    const lbl = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.id = id;
    cb.checked = xmlState.visibleKeys.includes(k);
    cb.addEventListener('change', () => {
      if (cb.checked) {
        if (!xmlState.visibleKeys.includes(k)) xmlState.visibleKeys.push(k);
      } else {
        xmlState.visibleKeys = xmlState.visibleKeys.filter((x) => x !== k);
        delete xmlState.filters[k];
      }
      renderTable();
    });
    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(' ' + k));
    wrap.appendChild(lbl);
  }
  $xmlColumns.replaceWith(wrap);
  wrap.id = 'xml-columns';
}

document.getElementById('xml-clear-filters').addEventListener('click', () => {
  xmlState.filters = {};
  renderTable();
});
document.getElementById('xml-clear-sort').addEventListener('click', () => {
  xmlState.sortBy = null;
  renderTable();
});

function compareValues(a, b) {
  const an = parseFloat(a);
  const bn = parseFloat(b);
  if (!isNaN(an) && !isNaN(bn) && /^-?\d+(\.\d+)?$/.test(a) && /^-?\d+(\.\d+)?$/.test(b)) {
    return an - bn;
  }
  return String(a).localeCompare(String(b));
}

function getFilteredSorted() {
  let rows = xmlState.rows;
  const filters = Object.entries(xmlState.filters).filter(([, v]) => v && v.length > 0);
  if (filters.length > 0) {
    rows = rows.filter((r) =>
      filters.every(([k, v]) => (r[k] || '').toLowerCase().includes(v.toLowerCase()))
    );
  }
  if (xmlState.sortBy) {
    const k = xmlState.sortBy;
    const dir = xmlState.sortDir === 'asc' ? 1 : -1;
    rows = [...rows].sort((a, b) => dir * compareValues(a[k] || '', b[k] || ''));
  }
  return rows;
}

function renderTable() {
  const thead = $xmlTable.querySelector('thead');
  const tbody = $xmlTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  const cols = xmlState.visibleKeys;
  if (cols.length === 0) return;

  // Header row with sort
  const hr = document.createElement('tr');
  for (const k of cols) {
    const th = document.createElement('th');
    const ind =
      xmlState.sortBy === k ? (xmlState.sortDir === 'asc' ? '▲' : '▼') : '';
    th.innerHTML = `${k} <span class="sort-ind">${ind}</span>`;
    th.addEventListener('click', () => {
      if (xmlState.sortBy === k) {
        xmlState.sortDir = xmlState.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        xmlState.sortBy = k;
        xmlState.sortDir = 'asc';
      }
      renderTable();
    });
    hr.appendChild(th);
  }
  thead.appendChild(hr);

  // Filter row
  const fr = document.createElement('tr');
  fr.className = 'filter-row';
  for (const k of cols) {
    const td = document.createElement('th');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = '필터…';
    inp.value = xmlState.filters[k] || '';
    inp.addEventListener('input', () => {
      xmlState.filters[k] = inp.value;
      renderTableBody();
    });
    inp.addEventListener('click', (e) => e.stopPropagation());
    td.appendChild(inp);
    fr.appendChild(td);
  }
  thead.appendChild(fr);

  renderTableBody();
}

function renderTableBody() {
  const tbody = $xmlTable.querySelector('tbody');
  tbody.innerHTML = '';
  const cols = xmlState.visibleKeys;
  const rows = getFilteredSorted();
  // Cap rendering to avoid huge DOM
  const MAX = 2000;
  const display = rows.slice(0, MAX);
  const frag = document.createDocumentFragment();
  for (const r of display) {
    const tr = document.createElement('tr');
    for (const k of cols) {
      const td = document.createElement('td');
      td.textContent = r[k] || '';
      td.title = r[k] || '';
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
  $xmlCount.textContent =
    `${rows.length}개 결과 (전체 ${xmlState.rows.length})` +
    (rows.length > MAX ? ` — 상위 ${MAX}개만 표시` : '');
}
