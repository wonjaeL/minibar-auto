// Headless verification: exercise the same logic as main.js IPC handlers.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');
const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
const pad4 = (s) => { s = String(s).trim(); return /^\d{3}$/.test(s) ? '0' + s : s; };

function getDayData(filePath, day) {
  const wb = XLSX.readFile(filePath);
  const dayStr = String(day);
  let matching = wb.SheetNames.filter((n) => n === dayStr);
  if (matching.length === 0) {
    const re = new RegExp(`^${dayStr}(\\s|\\(|$)`);
    matching = wb.SheetNames.filter((n) => re.test(n));
  }
  const dm = path.basename(filePath).match(/(\d{4})[_-](\d{2})/);
  let businessDate = null;
  if (dm) businessDate = `${String(day).padStart(2,'0')}-${MONTHS[parseInt(dm[2],10)-1]}-${dm[1].slice(-2)}`;

  const rooms = [], itemOrder = [], itemMap = new Map(), grid = {};
  for (const sn of matching) {
    const ws = wb.Sheets[sn];
    if (!ws['!ref']) continue;
    const ref = XLSX.utils.decode_range(ws['!ref']);
    const sheetRooms = [];
    for (let c = 7; c <= ref.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 4, c })];
      if (!cell || cell.v == null || cell.v === '') continue;
      let s = typeof cell.v === 'number' ? String(Math.trunc(cell.v)) : String(cell.v).trim();
      if (!/^\d+$/.test(s)) continue;
      s = pad4(s);
      sheetRooms.push({ col: c, room: s });
      if (!rooms.includes(s)) rooms.push(s);
      if (!grid[s]) grid[s] = {};
    }
    for (let r = 5; r <= ref.e.r; r++) {
      const nameCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
      if (!nameCell || nameCell.v == null || nameCell.v === '') break;
      const raw = String(nameCell.v).trim();
      if (/^total$/i.test(raw)) break;
      const name = raw.split('\n')[0].trim();
      if (!name) continue;
      const priceCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];
      const price = priceCell && typeof priceCell.v === 'number' ? priceCell.v : null;
      if (!itemMap.has(name)) { itemMap.set(name, { name, price }); itemOrder.push(name); }
      for (const { col, room } of sheetRooms) {
        const qCell = ws[XLSX.utils.encode_cell({ r, c: col })];
        if (!qCell) continue;
        const q = typeof qCell.v === 'number' ? qCell.v : parseInt(qCell.v, 10);
        if (!q || isNaN(q)) continue;
        grid[room][name] = (grid[room][name] || 0) + q;
      }
    }
  }
  return { rooms, items: itemOrder.map(n=>itemMap.get(n)), grid, businessDate, sheets: matching };
}

function collectRecords(node, out) {
  if (node == null) return;
  if (Array.isArray(node)) { for (const i of node) collectRecords(i, out); return; }
  if (typeof node !== 'object') return;
  for (const [k, v] of Object.entries(node)) {
    if (k === 'G_TRX_CHAR_DATE') { Array.isArray(v) ? out.push(...v) : out.push(v); }
    else collectRecords(v, out);
  }
}

function parseXml(filePath) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const parser = new XMLParser({ ignoreAttributes: true, trimValues: true, parseTagValue: false });
  const records = [];
  collectRecords(parser.parse(xml), records);
  return records.map((r) => Object.fromEntries(Object.entries(r).map(([k,v]) =>
    [k, v == null ? '' : (typeof v === 'object' ? JSON.stringify(v) : String(v).trim())])));
}

function matchXml(dayData, xmlRecords) {
  const xmlIndex = {};
  let xmlOnDateCount = 0;
  const dates = new Set();
  for (const r of xmlRecords) {
    if (r.BUSINESS_DATE) dates.add(r.BUSINESS_DATE);
    if (dayData.businessDate && r.BUSINESS_DATE !== dayData.businessDate) continue;
    xmlOnDateCount++;
    const room = pad4(r.ROOM || '');
    const item = String(r.ARTICLE_DESCRIPTION || '').trim();
    const qty = parseInt(r.QUANTITY, 10) || 0;
    if (!room || !item) continue;
    if (!xmlIndex[room]) xmlIndex[room] = {};
    xmlIndex[room][item] = (xmlIndex[room][item] || 0) + qty;
  }
  const stats = { match:0, qtyMismatch:0, onlyExcel:0, onlyXml:0 };
  const visited = new Set();
  for (const room of dayData.rooms) for (const it of dayData.items) {
    const eq = (dayData.grid[room]||{})[it.name] || 0;
    const xq = (xmlIndex[room]||{})[it.name] || 0;
    visited.add(`${room}|${it.name}`);
    if (eq===0 && xq===0) continue;
    if (eq===xq) stats.match++;
    else if (xq===0) stats.onlyExcel++;
    else if (eq===0) stats.onlyXml++;
    else stats.qtyMismatch++;
  }
  const extras = [];
  for (const room of Object.keys(xmlIndex)) for (const item of Object.keys(xmlIndex[room])) {
    if (!visited.has(`${room}|${item}`)) { extras.push({ room, item, xmlQty: xmlIndex[room][item] }); stats.onlyXml++; }
  }
  return { stats, extras, xmlOnDateCount, availableXmlDates: [...dates].sort() };
}

const xlsx = '7.Daily Minibar Sold_2026_04.xlsx';
const xml = 'finjrnl_articles_72599693.XML';

console.log('=== Apr 28 day data ===');
const d = getDayData(xlsx, 28);
console.log('rooms:', d.rooms.length, '| items:', d.items.length, '| businessDate:', d.businessDate);
console.log('items:', d.items.map(i => `${i.name}(${i.price})`).join(', '));
let totalQty=0, totalRev=0;
const priceMap = Object.fromEntries(d.items.map(i=>[i.name, i.price||0]));
for (const room of d.rooms) for (const [k,q] of Object.entries(d.grid[room]||{})) { totalQty+=q; totalRev+=q*priceMap[k]; }
console.log('total qty:', totalQty, '| revenue:', totalRev.toLocaleString());
// sample 3 rooms
for (const r of d.rooms.slice(0,3)) console.log('  ', r, d.grid[r]);

console.log('\n=== match Apr 28 vs XML (XML has May 03 data) ===');
const xr = parseXml(xml);
const m = matchXml(d, xr);
console.log('stats:', m.stats);
console.log('xmlOnDateCount:', m.xmlOnDateCount);
console.log('availableXmlDates:', m.availableXmlDates);
console.log('extras count:', m.extras.length);

console.log('\n=== match XML date directly: 03-MAY-26 ===');
const fakeDayData = { ...d, businessDate: '03-MAY-26' };
const m2 = matchXml(fakeDayData, xr);
console.log('stats:', m2.stats, '| xmlOnDateCount:', m2.xmlOnDateCount, '| extras:', m2.extras.length);
console.log('extras sample:', m2.extras.slice(0,5));
