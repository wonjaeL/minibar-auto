// ---------- Tabs ----------
document.querySelectorAll('.tab').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((b) => b.classList.remove('active'));
    document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
  });
});

// ---------- Shared state ----------
const app = {
  excelPath: null,
  dayData: null,    // { rooms, items, grid, businessDate, csv, sheets }
  matchResult: null,
  xml: { path: null, keys: [], rows: [] },
};

// ---------- Excel tab ----------
const $excelPath = document.getElementById('excel-path');
const $excelDate = document.getElementById('excel-date');
const $excelRun = document.getElementById('excel-run');
const $excelCopy = document.getElementById('excel-copy');
const $excelMatch = document.getElementById('excel-match');
const $excelOut = document.getElementById('excel-out');
const $excelMeta = document.getElementById('excel-meta');
const $excelSummary = document.getElementById('excel-summary');
const $excelBars = document.getElementById('excel-bars');
const $excelPivot = document.getElementById('excel-pivot');
const $excelExtras = document.getElementById('excel-extras');
const $legend = document.getElementById('legend');

function refreshExcelButtons() {
  $excelRun.disabled = !(app.excelPath && $excelDate.value);
  $excelCopy.disabled = !(app.dayData && app.dayData.rooms && app.dayData.rooms.length);
  $excelMatch.disabled = !(app.dayData && app.xml.rows.length);
  $excelMatch.title = !app.xml.rows.length
    ? 'XML Viewer 탭에서 XML을 먼저 로드하세요'
    : '';
}

document.getElementById('excel-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('excel');
  if (!p) return;
  app.excelPath = p;
  $excelPath.textContent = p;
  $excelPath.classList.remove('muted');
  refreshExcelButtons();
});

$excelDate.addEventListener('change', refreshExcelButtons);

$excelRun.addEventListener('click', async () => {
  const day = parseInt($excelDate.value.split('-')[2], 10);
  $excelMeta.textContent = '추출 중…';
  const r = await window.api.getDayData(app.excelPath, day);
  if (r.error) {
    $excelMeta.textContent = `오류: ${r.error}` + (r.sheets ? ` | 시트: ${r.sheets.join(', ')}` : '');
    app.dayData = null;
    refreshExcelButtons();
    return;
  }
  app.dayData = r;
  app.matchResult = null;
  $excelOut.value = r.csv;
  $excelMeta.textContent = `시트: ${r.sheets.join(', ')} · BUSINESS_DATE 후보: ${r.businessDate || '?'} · ${r.rooms.length} rooms · ${r.items.length} items`;
  renderSummary();
  renderBars();
  renderPivot();
  $excelExtras.innerHTML = '';
  $legend.classList.add('hidden');
  refreshExcelButtons();
});

$excelCopy.addEventListener('click', async () => {
  await navigator.clipboard.writeText($excelOut.value);
  $excelMeta.textContent = '클립보드에 복사됨';
});

$excelMatch.addEventListener('click', async () => {
  if (!app.dayData || !app.xml.rows.length) return;
  $excelMeta.textContent = 'XML 매칭 중…';
  const m = await window.api.matchXml(app.dayData, app.xml.rows);
  app.matchResult = m;
  let dateNote = '';
  if (m.xmlOnDateCount === 0 && m.targetDate) {
    dateNote = ` ⚠ XML에 ${m.targetDate} 데이터 없음 (있는 날짜: ${m.availableXmlDates.slice(0, 8).join(', ')}${m.availableXmlDates.length > 8 ? '…' : ''})`;
  }
  $excelMeta.textContent =
    `매칭(${m.targetDate || '?'}): ✅${m.stats.match} ⚠${m.stats.qtyMismatch} ❌Excel만 ${m.stats.onlyExcel} 🆕XML만 ${m.stats.onlyXml} (XML 레코드 ${m.xmlOnDateCount}개)${dateNote}`;
  $legend.classList.remove('hidden');
  renderSummary();
  renderPivot();
  renderExtras();
});

function renderSummary() {
  const d = app.dayData;
  if (!d) { $excelSummary.innerHTML = ''; return; }
  let totalQty = 0;
  let totalRevenue = 0;
  const priceMap = Object.fromEntries(d.items.map((i) => [i.name, i.price || 0]));
  for (const room of d.rooms) {
    for (const [item, q] of Object.entries(d.grid[room] || {})) {
      totalQty += q;
      totalRevenue += q * (priceMap[item] || 0);
    }
  }
  const cards = [
    { l: '날짜', v: d.businessDate || '-' },
    { l: 'Rooms', v: d.rooms.length },
    { l: 'Items', v: d.items.length },
    { l: '총 수량', v: totalQty },
    { l: '추정 매출', v: totalRevenue.toLocaleString() + '원' },
  ];
  if (app.matchResult) {
    const s = app.matchResult.stats;
    cards.push(
      { l: '매칭됨', v: s.match },
      { l: '수량 불일치', v: s.qtyMismatch },
      { l: 'Excel만', v: s.onlyExcel },
      { l: 'XML만', v: s.onlyXml },
    );
  }
  $excelSummary.innerHTML = cards
    .map((c) => `<div class="card"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`)
    .join('');
}

function renderBars() {
  const d = app.dayData;
  if (!d) { $excelBars.innerHTML = ''; return; }
  // sum qty per item
  const totals = {};
  for (const room of d.rooms) {
    for (const [it, q] of Object.entries(d.grid[room] || {})) {
      totals[it] = (totals[it] || 0) + q;
    }
  }
  const arr = d.items.map((i) => ({ name: i.name, qty: totals[i.name] || 0 }));
  arr.sort((a, b) => b.qty - a.qty);
  const max = Math.max(1, ...arr.map((x) => x.qty));
  $excelBars.innerHTML = arr
    .map(
      (x) => `
        <div class="blbl" title="${escapeHtml(x.name)}">${escapeHtml(x.name)}</div>
        <div class="btrack"><div class="bfill" style="width:${(x.qty / max) * 100}%"></div></div>
        <div class="bval">${x.qty}</div>`
    )
    .join('');
}

function renderPivot() {
  const d = app.dayData;
  const thead = $excelPivot.querySelector('thead');
  const tbody = $excelPivot.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  if (!d) return;

  // Index match status by room|item
  const matchIdx = {};
  if (app.matchResult) {
    for (const c of app.matchResult.cells) {
      matchIdx[`${c.room}|${c.item}`] = c.status;
    }
  }

  // Header
  const hr = document.createElement('tr');
  hr.innerHTML = `<th>Room \\ Item</th>` +
    d.items.map((i) => `<th title="${escapeHtml(i.name)} (${i.price || '-'}원)">${escapeHtml(i.name)}</th>`).join('') +
    `<th>합계</th>`;
  thead.appendChild(hr);

  // Body
  const frag = document.createDocumentFragment();
  for (const room of d.rooms) {
    const tr = document.createElement('tr');
    let rowSum = 0;
    let html = `<th>${room}</th>`;
    for (const it of d.items) {
      const q = (d.grid[room] && d.grid[room][it.name]) || 0;
      rowSum += q;
      const status = matchIdx[`${room}|${it.name}`];
      let cls = 'cell-empty';
      if (status) cls = 'cell-' + status;
      else if (q > 0) cls = 'cell-have';
      const txt = q ? q : '';
      const tip = app.matchResult
        ? `Excel:${q} / XML:${(app.matchResult.cells.find(c=>c.room===room && c.item===it.name) || {}).xmlQty || 0}`
        : '';
      html += `<td class="${cls}" title="${tip}">${txt}</td>`;
    }
    html += `<td><b>${rowSum || ''}</b></td>`;
    tr.innerHTML = html;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function renderExtras() {
  const m = app.matchResult;
  if (!m || !m.extras || m.extras.length === 0) { $excelExtras.innerHTML = ''; return; }
  const rows = m.extras
    .map((e) => `<tr><td>${e.room}</td><td>${escapeHtml(e.item)}</td><td>${e.xmlQty}</td></tr>`)
    .join('');
  $excelExtras.innerHTML = `
    <h4>XML에만 있는 항목 (Excel 그리드 외) — ${m.extras.length}건</h4>
    <table><thead><tr><th>Room</th><th>Item</th><th>XML Qty</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

// ---------- XML tab ----------
const xmlState = {
  visibleKeys: [],
  filters: {},
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
  app.xml.path = p;
  $xmlPath.textContent = p;
  $xmlPath.classList.remove('muted');
  $xmlCount.textContent = '파싱 중…';
  const r = await window.api.parseXml(p);
  if (r.error) { $xmlCount.textContent = `오류: ${r.error}`; return; }
  app.xml.keys = r.keys;
  app.xml.rows = r.rows;
  xmlState.visibleKeys = DEFAULT_COLUMNS.filter((k) => r.keys.includes(k));
  if (xmlState.visibleKeys.length === 0) xmlState.visibleKeys = r.keys.slice(0, 10);
  xmlState.filters = {};
  xmlState.sortBy = null;
  $xmlCount.textContent = `${r.count}개 레코드`;
  renderColumnPicker();
  renderXmlTable();
  refreshExcelButtons();
});

function renderColumnPicker() {
  $xmlColumns.className = 'col-picker';
  $xmlColumns.innerHTML = '';
  for (const k of app.xml.keys) {
    const lbl = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = xmlState.visibleKeys.includes(k);
    cb.addEventListener('change', () => {
      if (cb.checked) {
        if (!xmlState.visibleKeys.includes(k)) xmlState.visibleKeys.push(k);
      } else {
        xmlState.visibleKeys = xmlState.visibleKeys.filter((x) => x !== k);
        delete xmlState.filters[k];
      }
      renderXmlTable();
    });
    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(' ' + k));
    $xmlColumns.appendChild(lbl);
  }
}

document.getElementById('xml-clear-filters').addEventListener('click', () => { xmlState.filters = {}; renderXmlTable(); });
document.getElementById('xml-clear-sort').addEventListener('click', () => { xmlState.sortBy = null; renderXmlTable(); });

function compareValues(a, b) {
  if (/^-?\d+(\.\d+)?$/.test(a) && /^-?\d+(\.\d+)?$/.test(b)) return parseFloat(a) - parseFloat(b);
  return String(a).localeCompare(String(b));
}

function getFilteredSorted() {
  let rows = app.xml.rows;
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

function renderXmlTable() {
  const thead = $xmlTable.querySelector('thead');
  const tbody = $xmlTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';
  const cols = xmlState.visibleKeys;
  if (cols.length === 0) return;

  const hr = document.createElement('tr');
  for (const k of cols) {
    const th = document.createElement('th');
    const ind = xmlState.sortBy === k ? (xmlState.sortDir === 'asc' ? '▲' : '▼') : '';
    th.innerHTML = `${k} <span class="sort-ind">${ind}</span>`;
    th.addEventListener('click', () => {
      if (xmlState.sortBy === k) xmlState.sortDir = xmlState.sortDir === 'asc' ? 'desc' : 'asc';
      else { xmlState.sortBy = k; xmlState.sortDir = 'asc'; }
      renderXmlTable();
    });
    hr.appendChild(th);
  }
  thead.appendChild(hr);

  const fr = document.createElement('tr');
  fr.className = 'filter-row';
  for (const k of cols) {
    const td = document.createElement('th');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = '필터…';
    inp.value = xmlState.filters[k] || '';
    inp.addEventListener('input', () => { xmlState.filters[k] = inp.value; renderXmlBody(); });
    inp.addEventListener('click', (e) => e.stopPropagation());
    td.appendChild(inp);
    fr.appendChild(td);
  }
  thead.appendChild(fr);
  renderXmlBody();
}

function renderXmlBody() {
  const tbody = $xmlTable.querySelector('tbody');
  tbody.innerHTML = '';
  const cols = xmlState.visibleKeys;
  const rows = getFilteredSorted();
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
    `${rows.length}개 결과 (전체 ${app.xml.rows.length})` +
    (rows.length > MAX ? ` — 상위 ${MAX}개만 표시` : '');
}
