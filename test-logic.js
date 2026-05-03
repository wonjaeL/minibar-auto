// Headless verification of the same logic main.js uses (no Electron).
const fs = require('fs');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

// --- Excel: room list for a day ---
function getRooms(filePath, day) {
  const wb = XLSX.readFile(filePath);
  const dayStr = String(day);
  let matching = wb.SheetNames.filter((n) => n === dayStr);
  if (matching.length === 0) {
    const re = new RegExp(`^${dayStr}(\\s|\\(|$)`);
    matching = wb.SheetNames.filter((n) => re.test(n));
  }
  const rooms = [];
  for (const name of matching) {
    const ws = wb.Sheets[name];
    if (!ws['!ref']) continue;
    const ref = XLSX.utils.decode_range(ws['!ref']);
    for (let c = 7; c <= ref.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 4, c });
      const cell = ws[addr];
      if (!cell || cell.v === null || cell.v === undefined || cell.v === '') continue;
      let s = typeof cell.v === 'number' ? String(Math.trunc(cell.v)) : String(cell.v).trim();
      if (!/^\d+$/.test(s)) continue;
      if (s.length === 3) s = '0' + s;
      rooms.push(s);
    }
  }
  return { sheets: matching, csv: rooms.join(','), count: rooms.length };
}

// --- XML: collect all G_TRX_CHAR_DATE ---
function collectRecords(node, out) {
  if (node === null || node === undefined) return;
  if (Array.isArray(node)) { for (const i of node) collectRecords(i, out); return; }
  if (typeof node !== 'object') return;
  for (const [k, v] of Object.entries(node)) {
    if (k === 'G_TRX_CHAR_DATE') {
      if (Array.isArray(v)) out.push(...v); else out.push(v);
    } else collectRecords(v, out);
  }
}

function parseXml(filePath) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const parser = new XMLParser({ ignoreAttributes: true, trimValues: true, parseTagValue: false });
  const obj = parser.parse(xml);
  const records = [];
  collectRecords(obj, records);
  const keys = [];
  const seen = new Set();
  for (const r of records) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); keys.push(k); }
  return { keys, count: records.length, sample: records[0] };
}

// --- run ---
const xlsx = '7.Daily Minibar Sold_2026_04.xlsx';
const xml = 'finjrnl_articles_72599693.XML';

console.log('=== Excel Apr 28 ===');
console.log(getRooms(xlsx, 28));

console.log('\n=== Excel Apr 17 (split sheet) ===');
console.log(getRooms(xlsx, 17));

console.log('\n=== XML ===');
const r = parseXml(xml);
console.log('records:', r.count, '| keys:', r.keys.length);
console.log('sample row[0] subset:', {
  BUSINESS_DATE: r.sample.BUSINESS_DATE,
  BUSINESS_TIME: r.sample.BUSINESS_TIME,
  ROOM: r.sample.ROOM,
  ARTICLE_DESCRIPTION: r.sample.ARTICLE_DESCRIPTION,
  QUANTITY: r.sample.QUANTITY,
  GUEST_FULL_NAME: r.sample.GUEST_FULL_NAME,
});
