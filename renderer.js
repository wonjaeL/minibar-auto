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
  dayData: null,
  matchResult: null,
  xml: { path: null, keys: [], rows: [] },
  resv: { path: null, byRoom: {}, count: 0 },
  reportSort: { col: 'status', dir: 'asc' },
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

const $resvPath = document.getElementById('resv-path');
const $resvMeta = document.getElementById('resv-meta');
document.getElementById('resv-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('csv');
  if (!p) return;
  $resvPath.textContent = p;
  $resvPath.classList.remove('muted');
  $resvMeta.textContent = '파싱 중…';
  const r = await window.api.parseReservations(p);
  if (r.error) { $resvMeta.textContent = '오류: ' + r.error; return; }
  app.resv = { path: p, byRoom: r.byRoom, count: r.count };
  $resvMeta.textContent = `${r.count}건 예약 / ${Object.keys(r.byRoom).length} rooms`;
  if (app.matchResult) renderReport();
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
  renderReport();
});

// ---------- Report ----------
const $reportSection = document.getElementById('report-section');
const $reportTable = document.getElementById('report-table');
const $reportStatus = document.getElementById('report-status-filter');
const $reportSearch = document.getElementById('report-search');
$reportStatus.addEventListener('change', renderReport);
$reportSearch.addEventListener('input', renderReport);
document.getElementById('report-csv').addEventListener('click', exportReportCsv);

const REPORT_COLS = [
  { key: 'displayColor', label: 'Display' },
  { key: 'room', label: 'Room' },
  { key: 'confNo', label: 'Conf No' },
  { key: 'guest', label: 'Guest' },
  { key: 'item', label: 'Item' },
  { key: 'excelQty', label: 'Excel' },
  { key: 'xmlQty', label: 'XML' },
  { key: 'status', label: 'Status' },
];

function buildReportRows() {
  const m = app.matchResult;
  if (!m) return [];
  const rows = [];
  const lookupResv = (room) => {
    const list = app.resv.byRoom[room] || [];
    if (list.length === 0) return null;
    return list[0];
  };
  for (const c of m.cells) {
    if (c.status === 'empty') continue;
    const resv = lookupResv(c.room);
    rows.push({
      displayColor: resv?.displayColor || '',
      room: c.room,
      confNo: resv?.confirmationNumber || '',
      guest: resv?.name || (c.xmlGuests && c.xmlGuests[0]) || '',
      xmlGuestExtra: (c.xmlGuests && c.xmlGuests.length) ? c.xmlGuests.join(' / ') : '',
      item: c.item,
      excelQty: c.excelQty,
      xmlQty: c.xmlQty,
      status: c.status,
      _src: 'cell',
    });
  }
  for (const e of m.extras) {
    const resv = lookupResv(e.room);
    rows.push({
      displayColor: resv?.displayColor || '',
      room: e.room,
      confNo: resv?.confirmationNumber || '',
      guest: resv?.name || (e.xmlGuests && e.xmlGuests[0]) || '',
      xmlGuestExtra: (e.xmlGuests && e.xmlGuests.length) ? e.xmlGuests.join(' / ') : '',
      item: e.item,
      excelQty: 0,
      xmlQty: e.xmlQty,
      status: 'only-xml',
      _src: 'extra',
    });
  }
  return rows;
}

const STATUS_PRIO = { 'qty-mismatch': 0, 'only-excel': 1, 'only-xml': 2, 'match': 3 };

function filteredSortedReport() {
  let rows = buildReportRows();
  const sf = $reportStatus.value;
  if (sf === 'problems') rows = rows.filter(r => r.status !== 'match');
  else if (sf) rows = rows.filter(r => r.status === sf);
  const q = $reportSearch.value.trim().toLowerCase();
  if (q) rows = rows.filter(r =>
    [r.room, r.confNo, r.guest, r.item, r.xmlGuestExtra].some(v => String(v).toLowerCase().includes(q))
  );
  const { col, dir } = app.reportSort;
  const sign = dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let av = a[col], bv = b[col];
    if (col === 'status') { av = STATUS_PRIO[av] ?? 99; bv = STATUS_PRIO[bv] ?? 99; }
    if (typeof av === 'number' && typeof bv === 'number') return sign * (av - bv);
    return sign * String(av).localeCompare(String(bv));
  });
  return rows;
}

function renderReport() {
  if (!app.matchResult) { $reportSection.classList.add('hidden'); return; }
  $reportSection.classList.remove('hidden');
  const thead = $reportTable.querySelector('thead');
  const tbody = $reportTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const hr = document.createElement('tr');
  for (const c of REPORT_COLS) {
    const th = document.createElement('th');
    const ind = app.reportSort.col === c.key ? (app.reportSort.dir === 'asc' ? '▲' : '▼') : '';
    th.innerHTML = `${c.label} <span class="sort-ind">${ind}</span>`;
    th.addEventListener('click', () => {
      if (app.reportSort.col === c.key) app.reportSort.dir = app.reportSort.dir === 'asc' ? 'desc' : 'asc';
      else { app.reportSort.col = c.key; app.reportSort.dir = 'asc'; }
      renderReport();
    });
    hr.appendChild(th);
  }
  thead.appendChild(hr);

  const rows = filteredSortedReport();
  const frag = document.createDocumentFragment();
  for (const r of rows) {
    const tr = document.createElement('tr');
    tr.className = 'cell-' + r.status;
    const dc = r.displayColor;
    const dcCell = dc
      ? `<span class="display-color" style="background:${escapeHtml(dc)}" title="${escapeHtml(dc)}"></span>`
      : '';
    const guestTip = r.xmlGuestExtra && r.xmlGuestExtra !== r.guest
      ? ` (XML: ${escapeHtml(r.xmlGuestExtra)})` : '';
    tr.innerHTML =
      `<td>${dcCell}</td>` +
      `<td>${r.room}</td>` +
      `<td>${escapeHtml(r.confNo)}</td>` +
      `<td title="${escapeHtml(r.guest + guestTip)}">${escapeHtml(r.guest)}${guestTip ? ' *' : ''}</td>` +
      `<td>${escapeHtml(r.item)}</td>` +
      `<td>${r.excelQty || ''}</td>` +
      `<td>${r.xmlQty || ''}</td>` +
      `<td><b>${r.status}</b></td>`;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);

  const total = rows.length;
  const meta = document.getElementById('excel-meta');
  if (meta && app.matchResult) {
    // append/replace report count badge
    // (left as-is; cards already show stats)
  }
}

function exportReportCsv() {
  const rows = filteredSortedReport();
  const head = ['DisplayColor', 'Room', 'ConfNo', 'Guest', 'XMLGuests', 'Item', 'ExcelQty', 'XMLQty', 'Status'];
  const csvLines = [head.join(',')];
  for (const r of rows) {
    const cells = [r.displayColor, r.room, r.confNo, r.guest, r.xmlGuestExtra, r.item, r.excelQty, r.xmlQty, r.status]
      .map(v => {
        const s = String(v ?? '');
        return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
      });
    csvLines.push(cells.join(','));
  }
  const blob = new Blob([csvLines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const d = app.dayData?.businessDate || 'report';
  a.download = `minibar-report-${d}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

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

  // Header: items become rows; first col = Item, then each room
  const hr = document.createElement('tr');
  hr.innerHTML = `<th>Item \\ Room</th>` +
    d.rooms.map((r) => `<th>${r}</th>`).join('') +
    `<th>합계</th>`;
  thead.appendChild(hr);

  // Body: one row per item
  const frag = document.createDocumentFragment();
  for (const it of d.items) {
    const tr = document.createElement('tr');
    let rowSum = 0;
    let html = `<th title="${escapeHtml(it.name)} (${it.price || '-'}원)">${escapeHtml(it.name)}</th>`;
    for (const room of d.rooms) {
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
