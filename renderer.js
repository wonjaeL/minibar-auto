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
  resvDeparture: { path: null, byRoom: {}, count: 0 },
  resvInhouse:   { path: null, byRoom: {}, count: 0 },
  reportSort: { col: 'no', dir: 'asc' },
  actionSort: { col: 'priority', dir: 'asc' },
};

const $excelDate = document.getElementById('excel-date');
const $excelRun = document.getElementById('excel-run');
const $excelCopy = document.getElementById('excel-copy');
const $excelOut = document.getElementById('excel-out');
const $excelMeta = document.getElementById('excel-meta');
const $excelSummary = document.getElementById('excel-summary');
const $excelBars = document.getElementById('excel-bars');
const $excelPivot = document.getElementById('excel-pivot');
const $excelExtras = document.getElementById('excel-extras');
const $legend = document.getElementById('legend');

// ---------- Pipeline file pickers ----------
function setStepStatus(stepId, done, metaText) {
  const el = document.getElementById(stepId);
  el.classList.toggle('done', !!done);
  const meta = el.querySelector('.step-meta');
  meta.textContent = metaText;
  meta.classList.toggle('muted', !done);
}

document.getElementById('excel-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('excel');
  if (!p) return;
  app.excelPath = p;
  setStepStatus('step-excel', true, p.split(/[\\/]/).pop());
  // Auto-fill date if filename indicates year-month
  const fn = p.split(/[\\/]/).pop();
  const m = fn.match(/(\d{4})[_-](\d{2})/);
  if (m && !$excelDate.value) {
    // default to first day of that month
    $excelDate.value = `${m[1]}-${m[2]}-01`;
  }
  refreshButtons();
});

async function pickReservation(kind, stepId, stateKey) {
  const p = await window.api.selectFile('csv');
  if (!p) return;
  setStepStatus(stepId, false, '파싱 중…');
  const r = await window.api.parseReservations(p);
  if (r.error) { setStepStatus(stepId, false, '오류: ' + r.error); return; }
  app[stateKey] = { path: p, byRoom: r.byRoom, count: r.count };
  const fn = p.split(/[\\/]/).pop();
  setStepStatus(stepId, true, `${fn} · ${r.count}건 / ${Object.keys(r.byRoom).length} rooms`);
  if (app.dayData) generateReport();
}

document.getElementById('dep-pick').addEventListener('click', () => pickReservation('departure', 'step-dep', 'resvDeparture'));
document.getElementById('inh-pick').addEventListener('click', () => pickReservation('inhouse', 'step-inh', 'resvInhouse'));

// XML picker (mirrored to step-4)
async function loadXml(path) {
  setStepStatus('step-xml', false, '파싱 중…');
  const r = await window.api.parseXml(path);
  if (r.error) { setStepStatus('step-xml', false, '오류: ' + r.error); return; }
  app.xml = { path, keys: r.keys, rows: r.rows };
  const fn = path.split(/[\\/]/).pop();
  setStepStatus('step-xml', true, `${fn} · ${r.count} records`);
  // refresh XML tab UI as well
  document.getElementById('xml-path').textContent = path;
  document.getElementById('xml-path').classList.remove('muted');
  document.getElementById('xml-count').textContent = `${r.count}개 레코드`;
  xmlState.visibleKeys = DEFAULT_COLUMNS.filter((k) => r.keys.includes(k));
  if (xmlState.visibleKeys.length === 0) xmlState.visibleKeys = r.keys.slice(0, 10);
  xmlState.filters = {};
  xmlState.sortBy = null;
  renderColumnPicker();
  renderXmlTable();
  if (app.dayData) generateReport();
}

document.getElementById('xml-pick-2').addEventListener('click', async () => {
  const p = await window.api.selectFile('xml');
  if (p) await loadXml(p);
});
document.getElementById('xml-pick').addEventListener('click', async () => {
  const p = await window.api.selectFile('xml');
  if (p) await loadXml(p);
});

$excelDate.addEventListener('change', refreshButtons);

function refreshButtons() {
  $excelRun.disabled = !(app.excelPath && $excelDate.value);
  $excelCopy.disabled = !(app.dayData && app.dayData.rooms && app.dayData.rooms.length);
}

// ---------- Generate report ----------
$excelRun.addEventListener('click', () => generateReport());

async function generateReport() {
  if (!app.excelPath || !$excelDate.value) return;
  const day = parseInt($excelDate.value.split('-')[2], 10);
  $excelMeta.textContent = '추출 중…';
  const r = await window.api.getDayData(app.excelPath, day);
  if (r.error) {
    $excelMeta.textContent = `오류: ${r.error}` + (r.sheets ? ` | 시트: ${r.sheets.join(', ')}` : '');
    app.dayData = null;
    refreshButtons();
    return;
  }
  app.dayData = r;
  $excelOut.value = r.csv;

  // If XML loaded, also do matching
  if (app.xml.rows.length) {
    const m = await window.api.matchXml(app.dayData, app.xml.rows);
    app.matchResult = m;
    $excelMeta.textContent =
      `${r.businessDate} · 시트 ${r.sheets.join(', ')} · ${r.rooms.length} rooms × ${r.items.length} items · XML ${m.xmlOnDateCount}건` +
      (m.xmlOnDateCount === 0 && m.targetDate
        ? ` ⚠ XML에 ${m.targetDate} 데이터 없음 (있는 날짜: ${m.availableXmlDates.slice(0,5).join(', ')}${m.availableXmlDates.length>5?'…':''})`
        : '');
    $legend.classList.remove('hidden');
  } else {
    app.matchResult = null;
    $excelMeta.textContent = `${r.businessDate} · ${r.rooms.length} rooms × ${r.items.length} items · XML 미로드`;
    $legend.classList.add('hidden');
  }

  renderSummary();
  renderBars();
  renderPivot();
  renderExtras();
  renderActionReport();
  renderReport();
  refreshButtons();
}

$excelCopy.addEventListener('click', async () => {
  await navigator.clipboard.writeText($excelOut.value);
  $excelMeta.textContent = '클립보드에 복사됨';
});

// ---------- Guest lookup ----------
function lookupGuest(room) {
  const dep = (app.resvDeparture.byRoom[room] || [])[0];
  if (dep) return { kind: 'departure', ...dep };
  const inh = (app.resvInhouse.byRoom[room] || [])[0];
  if (inh) return { kind: 'inhouse', ...inh };
  return { kind: 'unknown' };
}

const TITLE_RE = /,\s*(MR|MRS|MS|MISS|DR|MSTR|MASTER|SIR|MADAM)\.?\s*$/i;
function cleanGuestName(s) {
  if (!s) return '';
  let n = String(s).trim().replace(/\s*\(\d+\)\s*$/, '').trim();
  while (TITLE_RE.test(n)) n = n.replace(TITLE_RE, '').trim();
  return n;
}

// ---------- Cards / bars / pivot (existing) ----------
function renderSummary() {
  const d = app.dayData;
  if (!d) { $excelSummary.innerHTML = ''; return; }
  let totalQty = 0, totalRevenue = 0;
  const priceMap = Object.fromEntries(d.items.map((i) => [i.name, i.price || 0]));
  for (const room of d.rooms) {
    for (const [it, q] of Object.entries(d.grid[room] || {})) {
      totalQty += q;
      totalRevenue += q * (priceMap[it] || 0);
    }
  }
  const cards = [
    { l: '날짜', v: d.businessDate || '-' },
    { l: 'Rooms', v: d.rooms.length },
    { l: 'Items', v: d.items.length },
    { l: '총 수량', v: totalQty },
    { l: '소비 추정 매출', v: totalRevenue.toLocaleString() + '원' },
  ];
  if (app.matchResult) {
    const action = buildActionRows();
    let critical = 0, warning = 0, review = 0, totalLoss = 0;
    for (const r of action) {
      if (r.severity === 'critical') critical++;
      else if (r.severity === 'warning') warning++;
      else if (r.severity === 'review') review++;
      totalLoss += r.netLoss;
    }
    cards.push(
      { l: '🚨 체크아웃-미청구', v: critical },
      { l: '⚠️ 인하우스-미청구', v: warning },
      { l: 'ℹ️ 검토 필요', v: review },
      { l: '예상 순손실', v: totalLoss.toLocaleString() + '원' },
    );
  }
  $excelSummary.innerHTML = cards
    .map((c) => `<div class="card"><div class="v">${c.v}</div><div class="l">${c.l}</div></div>`)
    .join('');
}

function renderBars() {
  const d = app.dayData;
  if (!d) { $excelBars.innerHTML = ''; return; }
  const totals = {};
  for (const room of d.rooms) for (const [it, q] of Object.entries(d.grid[room] || {})) totals[it] = (totals[it] || 0) + q;
  const arr = d.items.map((i) => ({ name: i.name, qty: totals[i.name] || 0 }));
  arr.sort((a, b) => b.qty - a.qty);
  const max = Math.max(1, ...arr.map((x) => x.qty));
  $excelBars.innerHTML = arr
    .map((x) => `
      <div class="blbl" title="${escapeHtml(x.name)}">${escapeHtml(x.name)}</div>
      <div class="btrack"><div class="bfill" style="width:${(x.qty / max) * 100}%"></div></div>
      <div class="bval">${x.qty}</div>`)
    .join('');
}

function renderPivot() {
  const d = app.dayData;
  const thead = $excelPivot.querySelector('thead');
  const tbody = $excelPivot.querySelector('tbody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  if (!d) return;
  const matchIdx = {};
  if (app.matchResult) for (const c of app.matchResult.cells) matchIdx[`${c.room}|${c.item}`] = c.status;

  const hr = document.createElement('tr');
  hr.innerHTML = `<th>Item \\ Room</th>` + d.rooms.map((r) => `<th>${r}</th>`).join('') + `<th>합계</th>`;
  thead.appendChild(hr);

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
  const extrasBtn = document.getElementById('extras-copy');
  if (!m || !m.extras || m.extras.length === 0) {
    $excelExtras.innerHTML = '';
    if (extrasBtn) extrasBtn.style.display = 'none';
    return;
  }
  if (extrasBtn) extrasBtn.style.display = '';
  const rows = m.extras.map((e) => `<tr><td>${e.room}</td><td>${escapeHtml(e.item)}</td><td>${e.xmlQty}</td></tr>`).join('');
  $excelExtras.innerHTML = `
    <h4>XML에만 있는 항목 (Excel 그리드 외) — ${m.extras.length}건</h4>
    <table><thead><tr><th>Room</th><th>Item</th><th>XML Qty</th></tr></thead><tbody>${rows}</tbody></table>`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

// TSV (tab-separated) — Excel pastes this directly into cells
function toTSV(headers, rows) {
  const esc = (v) => {
    const s = v == null ? '' : String(v);
    return /[\t\n\r"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(esc).join('\t')];
  for (const r of rows) lines.push(r.map(esc).join('\t'));
  return lines.join('\r\n');
}

async function copyToClipboardTSV(headers, rows, btn) {
  try {
    await navigator.clipboard.writeText(toTSV(headers, rows));
    if (btn) {
      const orig = btn.textContent;
      btn.textContent = '✓ 복사됨';
      btn.disabled = true;
      setTimeout(() => { btn.textContent = orig; btn.disabled = false; }, 1500);
    }
  } catch (e) {
    alert('복사 실패: ' + e.message);
  }
}

// ---------- Action Report (NEW: priority-driven reconciliation) ----------
const $actionTitle = document.getElementById('action-title');
const $actionControls = document.getElementById('action-controls');
const $actionWrap = document.getElementById('action-table-wrap');
const $actionTable = document.getElementById('action-table');
const $actionFilter = document.getElementById('action-filter');
const $actionSearch = document.getElementById('action-search');
$actionFilter.addEventListener('change', renderActionReport);
$actionSearch.addEventListener('input', renderActionReport);
document.getElementById('action-csv').addEventListener('click', exportActionCsv);
document.getElementById('action-copy').addEventListener('click', (ev) => {
  const rows = filteredActionRows();
  const head = ['#', '유형', 'Room', 'Conf No', 'Name', '객실 상태', '항목', '예상 손실(원)'];
  const body = rows.map((r) => [r.priority, r.severityLabel, r.room, r.confNo, r.name, r.guestKind, r.issuesText, r.netLoss]);
  copyToClipboardTSV(head, body, ev.currentTarget);
});

const ACTION_COLS = [
  { key: 'priority', label: '#' },
  { key: 'severity', label: '유형' },
  { key: 'room', label: 'Room' },
  { key: 'confNo', label: 'Conf No' },
  { key: 'name', label: 'Name' },
  { key: 'guestKind', label: '객실 상태' },
  { key: 'issues', label: '항목' },
  { key: 'netLoss', label: '예상 손실(원)' },
];

// Per-cell classification for action purposes
function classifyItem(c) {
  if (c.excelQty === 0 && c.xmlQty === 0) return null;
  if (c.excelQty === c.xmlQty) return 'match';
  if (c.xmlQty === 0) return 'unbilled';
  if (c.excelQty === 0) return 'ghost';
  if (c.excelQty > c.xmlQty) return 'partial-unbilled';
  return 'overbilled';
}

function buildActionRows() {
  if (!app.matchResult) return [];
  const priceMap = Object.fromEntries((app.dayData?.items || []).map((i) => [i.name, i.price || 0]));
  const byRoom = {};
  const ensure = (r) => byRoom[r] = byRoom[r] || { room: r, issues: [], netLoss: 0, kinds: new Set() };

  for (const c of app.matchResult.cells) {
    const k = classifyItem(c);
    if (!k || k === 'match') continue;
    const e = ensure(c.room);
    const price = priceMap[c.item] || 0;
    let loss = 0, label = '';
    if (k === 'unbilled') {
      loss = price * c.excelQty;
      label = `${c.item} ${c.excelQty} 미청구`;
    } else if (k === 'partial-unbilled') {
      loss = price * (c.excelQty - c.xmlQty);
      label = `${c.item} ${c.excelQty - c.xmlQty} 부분미청구 (HK ${c.excelQty}/PMS ${c.xmlQty})`;
    } else if (k === 'overbilled') {
      loss = -price * (c.xmlQty - c.excelQty);
      label = `${c.item} ${c.xmlQty - c.excelQty} 과청구 (HK ${c.excelQty}/PMS ${c.xmlQty})`;
    } else if (k === 'ghost') {
      loss = -price * c.xmlQty;
      label = `${c.item} ${c.xmlQty} HK미보고 (PMS만)`;
    }
    e.issues.push({ kind: k, label, loss });
    e.netLoss += loss;
    e.kinds.add(k);
  }
  for (const x of app.matchResult.extras) {
    const e = ensure(x.room);
    const price = priceMap[x.item] || 0;
    const loss = -price * x.xmlQty;
    e.issues.push({ kind: 'ghost', label: `${x.item} ${x.xmlQty} HK미보고 (PMS만)`, loss });
    e.netLoss += loss;
    e.kinds.add('ghost');
  }

  const rows = Object.values(byRoom).map((r) => {
    const guest = lookupGuest(r.room);
    const hasUnbilled = r.kinds.has('unbilled') || r.kinds.has('partial-unbilled');
    let severity = 'review';
    let severityLabel = '검토';
    if (guest.kind === 'departure' && hasUnbilled) { severity = 'critical'; severityLabel = '체크아웃-미청구'; }
    else if (guest.kind === 'inhouse' && hasUnbilled) { severity = 'warning'; severityLabel = '인하우스-미청구'; }
    else if (hasUnbilled) { severity = 'warning'; severityLabel = '미청구 (예약불명)'; }
    else if (r.kinds.has('overbilled')) { severity = 'review'; severityLabel = '과청구'; }
    else if (r.kinds.has('ghost')) { severity = 'review'; severityLabel = 'HK 미보고'; }
    return {
      room: r.room,
      confNo: guest.confirmationNumber || '',
      name: cleanGuestName(guest.name),
      guestKind: guest.kind,
      issues: r.issues,
      issuesText: r.issues.map((i) => i.label).join(', '),
      netLoss: r.netLoss,
      severity, severityLabel,
    };
  });
  // Sort: critical → warning → review, then by net loss desc, then room asc
  const sevPrio = { critical: 0, warning: 1, review: 2 };
  rows.sort((a, b) =>
    (sevPrio[a.severity] - sevPrio[b.severity]) ||
    (b.netLoss - a.netLoss) ||
    String(a.room).localeCompare(String(b.room))
  );
  rows.forEach((r, i) => r.priority = i + 1);
  return rows;
}

function filteredActionRows() {
  let rows = buildActionRows();
  const f = $actionFilter.value;
  if (f) rows = rows.filter((r) => r.severity === f);
  const q = $actionSearch.value.trim().toLowerCase();
  if (q) rows = rows.filter((r) =>
    [r.room, r.confNo, r.name, r.issuesText].some((v) => String(v).toLowerCase().includes(q))
  );
  return rows;
}

function renderActionReport() {
  const rows = buildActionRows();
  if (!app.matchResult || rows.length === 0) {
    $actionTitle.style.display = 'none';
    $actionControls.style.display = 'none';
    $actionWrap.style.display = 'none';
    return;
  }
  $actionTitle.style.display = '';
  $actionControls.style.display = '';
  $actionWrap.style.display = '';

  const display = filteredActionRows();
  const thead = $actionTable.querySelector('thead');
  const tbody = $actionTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const hr = document.createElement('tr');
  for (const c of ACTION_COLS) {
    const th = document.createElement('th');
    th.textContent = c.label;
    hr.appendChild(th);
  }
  thead.appendChild(hr);

  const frag = document.createDocumentFragment();
  const guestKindLabel = { departure: '체크아웃', inhouse: '인하우스', unknown: '예약 없음' };
  for (const r of display) {
    const tr = document.createElement('tr');
    tr.className = 'sev-' + r.severity;
    const tag = `<span class="sev-tag ${r.severity}">${r.severity === 'critical' ? '🚨 ' : r.severity === 'warning' ? '⚠️ ' : 'ℹ️ '}${escapeHtml(r.severityLabel)}</span>`;
    const lossCls = r.netLoss > 0 ? 'neg' : (r.netLoss < 0 ? 'pos' : '');
    const lossStr = r.netLoss === 0 ? '0' : (r.netLoss > 0 ? `-${r.netLoss.toLocaleString()}` : `+${(-r.netLoss).toLocaleString()}`);
    tr.innerHTML =
      `<td>${r.priority}</td>` +
      `<td>${tag}</td>` +
      `<td>${r.room}</td>` +
      `<td>${escapeHtml(r.confNo)}</td>` +
      `<td>${escapeHtml(r.name)}</td>` +
      `<td><span class="guest-kind ${r.guestKind}">${guestKindLabel[r.guestKind]}</span></td>` +
      `<td>${escapeHtml(r.issuesText)}</td>` +
      `<td class="loss-cell ${lossCls}">${lossStr}</td>`;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function exportActionCsv() {
  const rows = filteredActionRows();
  const head = ['#', 'Severity', 'Room', 'ConfNo', 'Name', 'GuestStatus', 'Issues', 'NetLoss(원)'];
  const lines = [head.join(',')];
  for (const r of rows) {
    const cells = [r.priority, r.severityLabel, r.room, r.confNo, r.name, r.guestKind, r.issuesText, r.netLoss]
      .map((v) => {
        const s = String(v ?? '');
        return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
      });
    lines.push(cells.join(','));
  }
  const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `minibar-action-${app.dayData?.businessDate || 'report'}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------- Aggregated Report (existing 4-col, enhanced with kind) ----------
const $reportTitle = document.getElementById('report-title');
const $reportControls = document.getElementById('report-controls');
const $reportWrap = document.getElementById('report-table-wrap');
const $reportTable = document.getElementById('report-table');
const $reportStatus = document.getElementById('report-status-filter');
const $reportSearch = document.getElementById('report-search');
$reportStatus.addEventListener('change', renderReport);
$reportSearch.addEventListener('input', renderReport);
document.getElementById('report-csv').addEventListener('click', exportReportCsv);
document.getElementById('report-copy').addEventListener('click', (ev) => {
  const rows = filteredSortedReport();
  const head = ['No', 'Room', 'Conf No', 'Name', '상태', 'Minibar List'];
  const body = rows.map((r, i) => [i + 1, r.room, r.confNo, r.name, r.guestKind, r.minibar]);
  copyToClipboardTSV(head, body, ev.currentTarget);
});

// Pivot copy
document.getElementById('pivot-copy').addEventListener('click', (ev) => {
  const d = app.dayData;
  if (!d) return;
  const head = ['Item', ...d.rooms, '합계'];
  const body = d.items.map((it) => {
    let sum = 0;
    const cells = d.rooms.map((room) => {
      const q = (d.grid[room] && d.grid[room][it.name]) || 0;
      sum += q;
      return q || '';
    });
    return [it.name, ...cells, sum || ''];
  });
  copyToClipboardTSV(head, body, ev.currentTarget);
});

// Extras (XML-only) copy
document.getElementById('extras-copy').addEventListener('click', (ev) => {
  const m = app.matchResult;
  if (!m || !m.extras) return;
  const head = ['Room', 'Item', 'XML Qty'];
  const body = m.extras.map((e) => [e.room, e.item, e.xmlQty]);
  copyToClipboardTSV(head, body, ev.currentTarget);
});

// XML viewer copy
document.getElementById('xml-copy').addEventListener('click', (ev) => {
  const cols = xmlState.visibleKeys;
  const rows = getFilteredSorted();
  const body = rows.map((r) => cols.map((k) => r[k] || ''));
  copyToClipboardTSV(cols, body, ev.currentTarget);
});

const REPORT_COLS = [
  { key: 'no', label: 'No' },
  { key: 'room', label: 'Room' },
  { key: 'confNo', label: 'Conf No' },
  { key: 'name', label: 'Name' },
  { key: 'guestKind', label: '상태' },
  { key: 'minibar', label: 'Minibar List' },
];

function pickOverallStatus(statuses) {
  if (statuses.has('only-excel')) return 'only-excel';
  if (statuses.has('qty-mismatch')) return 'qty-mismatch';
  if (statuses.has('only-xml')) return 'only-xml';
  if (statuses.has('match')) return 'match';
  return '';
}

function formatMinibarList(items) {
  return items
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((it) => {
      const e = it.excelQty || 0, x = it.xmlQty || 0;
      if (e === x) return `${it.name} ${e}`;
      if (e > 0 && x === 0) return `${it.name} ${e}`;
      if (e === 0 && x > 0) return `${it.name} ${x}`;
      return `${it.name} ${e}/${x}`;
    })
    .join(', ');
}

function buildReportRows() {
  const m = app.matchResult;
  if (!m) return [];
  const byRoom = {};
  const ensure = (room) => byRoom[room] = byRoom[room] || { room, items: [], statuses: new Set() };
  for (const c of m.cells) {
    const e = ensure(c.room);
    if (c.excelQty > 0 || c.xmlQty > 0) e.items.push({ name: c.item, excelQty: c.excelQty, xmlQty: c.xmlQty, status: c.status });
    if (c.status && c.status !== 'empty') e.statuses.add(c.status);
  }
  for (const x of m.extras) {
    const e = ensure(x.room);
    e.items.push({ name: x.item, excelQty: 0, xmlQty: x.xmlQty, status: 'only-xml' });
    e.statuses.add('only-xml');
  }
  const rows = [];
  for (const room of Object.keys(byRoom)) {
    const r = byRoom[room];
    if (r.items.length === 0) continue;
    const guest = lookupGuest(room);
    rows.push({
      room,
      confNo: guest.confirmationNumber || '',
      name: cleanGuestName(guest.name),
      guestKind: guest.kind,
      items: r.items,
      minibar: formatMinibarList(r.items),
      status: pickOverallStatus(r.statuses),
    });
  }
  return rows;
}

const STATUS_PRIO = { 'qty-mismatch': 0, 'only-excel': 1, 'only-xml': 2, 'match': 3, '': 4 };

function filteredSortedReport() {
  let rows = buildReportRows();
  const sf = $reportStatus.value;
  if (sf === 'problems') rows = rows.filter((r) => r.status && r.status !== 'match');
  else if (sf) rows = rows.filter((r) => r.status === sf);
  const q = $reportSearch.value.trim().toLowerCase();
  if (q) rows = rows.filter((r) => [r.room, r.confNo, r.name, r.minibar].some((v) => String(v).toLowerCase().includes(q)));
  const { col, dir } = app.reportSort;
  const sign = dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    if (col === 'no' || col === 'room') return sign * String(a.room).localeCompare(String(b.room));
    if (col === 'minibar') return sign * (b.items.length - a.items.length);
    return sign * String(a[col] || '').localeCompare(String(b[col] || ''));
  });
  return rows;
}

function renderReport() {
  if (!app.matchResult) {
    $reportTitle.style.display = 'none';
    $reportControls.style.display = 'none';
    $reportWrap.style.display = 'none';
    return;
  }
  $reportTitle.style.display = '';
  $reportControls.style.display = '';
  $reportWrap.style.display = '';

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

  const guestLabel = { departure: '체크아웃', inhouse: '인하우스', unknown: '예약없음' };
  const rows = filteredSortedReport();
  const frag = document.createDocumentFragment();
  let i = 0;
  for (const r of rows) {
    i++;
    const tr = document.createElement('tr');
    tr.className = 'cell-' + (r.status || 'empty');
    tr.innerHTML =
      `<td>${i}</td>` +
      `<td>${r.room}</td>` +
      `<td>${escapeHtml(r.confNo)}</td>` +
      `<td>${escapeHtml(r.name)}</td>` +
      `<td><span class="guest-kind ${r.guestKind}">${guestLabel[r.guestKind] || r.guestKind}</span></td>` +
      `<td>${escapeHtml(r.minibar)}</td>`;
    frag.appendChild(tr);
  }
  tbody.appendChild(frag);
}

function exportReportCsv() {
  const rows = filteredSortedReport();
  const head = ['No', 'Room', 'ConfNo', 'Name', 'GuestStatus', 'MinibarList', 'OverallStatus'];
  const lines = [head.join(',')];
  let i = 0;
  for (const r of rows) {
    i++;
    const cells = [i, r.room, r.confNo, r.name, r.guestKind, r.minibar, r.status]
      .map((v) => { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; });
    lines.push(cells.join(','));
  }
  const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `minibar-report-${app.dayData?.businessDate || 'report'}.csv`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ---------- XML tab ----------
const xmlState = { visibleKeys: [], filters: {}, sortBy: null, sortDir: 'asc' };
const DEFAULT_COLUMNS = [
  'BUSINESS_DATE', 'BUSINESS_TIME', 'ROOM', 'ARTICLE_DESCRIPTION', 'QUANTITY',
  'CASHIER_DEBIT', 'CASHIER_CREDIT', 'GUEST_FULL_NAME', 'TRX_NO', 'TRX_DESC',
  'MARKET_CODE', 'ROOM_CLASS', 'CASHIER_NAME',
];
const $xmlColumns = document.getElementById('xml-columns');
const $xmlTable = document.getElementById('xml-table');
const $xmlCount = document.getElementById('xml-count');

function renderColumnPicker() {
  $xmlColumns.className = 'col-picker';
  $xmlColumns.innerHTML = '';
  for (const k of app.xml.keys) {
    const lbl = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = xmlState.visibleKeys.includes(k);
    cb.addEventListener('change', () => {
      if (cb.checked) { if (!xmlState.visibleKeys.includes(k)) xmlState.visibleKeys.push(k); }
      else { xmlState.visibleKeys = xmlState.visibleKeys.filter((x) => x !== k); delete xmlState.filters[k]; }
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
  if (filters.length > 0) rows = rows.filter((r) => filters.every(([k, v]) => (r[k] || '').toLowerCase().includes(v.toLowerCase())));
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
  thead.innerHTML = ''; tbody.innerHTML = '';
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
    inp.type = 'text'; inp.placeholder = '필터…';
    inp.value = xmlState.filters[k] || '';
    inp.addEventListener('input', () => { xmlState.filters[k] = inp.value; renderXmlBody(); });
    inp.addEventListener('click', (e) => e.stopPropagation());
    td.appendChild(inp); fr.appendChild(td);
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
  $xmlCount.textContent = `${rows.length}개 결과 (전체 ${app.xml.rows.length})` + (rows.length > MAX ? ` — 상위 ${MAX}개만 표시` : '');
}
