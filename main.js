const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

const MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];

function pad4(s) {
  s = String(s).trim();
  if (/^\d{3}$/.test(s)) return '0' + s;
  return s;
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 820,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile('index.html');
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('select-file', async (_evt, kind) => {
  const filters =
    kind === 'xml'
      ? [{ name: 'XML', extensions: ['xml', 'XML'] }]
      : [{ name: 'Excel', extensions: ['xlsx', 'xls'] }];
  const r = await dialog.showOpenDialog({
    title: kind === 'xml' ? 'Select finjrnl XML' : 'Select minibar Excel',
    filters,
    properties: ['openFile'],
  });
  if (r.canceled || r.filePaths.length === 0) return null;
  return r.filePaths[0];
});

ipcMain.handle('get-day-data', async (_evt, filePath, day) => {
  try {
    const wb = XLSX.readFile(filePath);
    const dayStr = String(day);
    let matching = wb.SheetNames.filter((n) => n === dayStr);
    if (matching.length === 0) {
      const re = new RegExp(`^${dayStr}(\\s|\\(|$)`);
      matching = wb.SheetNames.filter((n) => re.test(n));
    }
    if (matching.length === 0) {
      return { error: `Sheet for day ${day} not found`, sheets: wb.SheetNames };
    }

    const baseName = path.basename(filePath);
    const dm = baseName.match(/(\d{4})[_-](\d{2})/);
    let businessDate = null;
    if (dm) {
      const year = dm[1];
      const month = parseInt(dm[2], 10);
      businessDate = `${String(day).padStart(2, '0')}-${MONTHS[month - 1]}-${year.slice(-2)}`;
    }

    const rooms = [];
    const itemOrder = [];
    const itemMap = new Map(); // name → { name, price }
    const grid = {}; // room → { item → qty }

    for (const sn of matching) {
      const ws = wb.Sheets[sn];
      if (!ws['!ref']) continue;
      const ref = XLSX.utils.decode_range(ws['!ref']);

      // Row index 4 (sheet row 5) → room numbers in cols 7..end
      const sheetRooms = [];
      for (let c = 7; c <= ref.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: 4, c })];
        if (!cell || cell.v === null || cell.v === undefined || cell.v === '') continue;
        let s = typeof cell.v === 'number' ? String(Math.trunc(cell.v)) : String(cell.v).trim();
        if (!/^\d+$/.test(s)) continue;
        s = pad4(s);
        sheetRooms.push({ col: c, room: s });
        if (!rooms.includes(s)) rooms.push(s);
        if (!grid[s]) grid[s] = {};
      }

      // Item rows from row index 5 (sheet row 6) onward; stop at empty/TOTAL
      for (let r = 5; r <= ref.e.r; r++) {
        const nameCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
        if (!nameCell || nameCell.v === null || nameCell.v === undefined || nameCell.v === '') break;
        const raw = String(nameCell.v).trim();
        if (/^total$/i.test(raw)) break;
        const itemName = raw.split('\n')[0].trim();
        if (!itemName) continue;
        const priceCell = ws[XLSX.utils.encode_cell({ r, c: 4 })];
        const price = priceCell && typeof priceCell.v === 'number' ? priceCell.v : null;
        if (!itemMap.has(itemName)) {
          itemMap.set(itemName, { name: itemName, price });
          itemOrder.push(itemName);
        }
        for (const { col, room } of sheetRooms) {
          const qCell = ws[XLSX.utils.encode_cell({ r, c: col })];
          if (!qCell) continue;
          const q = typeof qCell.v === 'number' ? qCell.v : parseInt(qCell.v, 10);
          if (!q || isNaN(q)) continue;
          grid[room][itemName] = (grid[room][itemName] || 0) + q;
        }
      }
    }

    const items = itemOrder.map((n) => itemMap.get(n));
    return {
      rooms,
      items,
      grid,
      businessDate,
      sheets: matching,
      csv: rooms.join(','),
    };
  } catch (e) {
    return { error: e.message };
  }
});

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

ipcMain.handle('parse-xml', async (_evt, filePath) => {
  try {
    const xml = fs.readFileSync(filePath, 'utf8');
    const parser = new XMLParser({ ignoreAttributes: true, trimValues: true, parseTagValue: false });
    const obj = parser.parse(xml);
    const records = [];
    collectRecords(obj, records);
    const flat = records.map((r) => {
      const out = {};
      for (const [k, v] of Object.entries(r)) {
        if (v === null || v === undefined) out[k] = '';
        else if (typeof v === 'object') out[k] = JSON.stringify(v);
        else out[k] = String(v).trim();
      }
      return out;
    });
    const keys = [];
    const seen = new Set();
    for (const r of flat) for (const k of Object.keys(r)) if (!seen.has(k)) { seen.add(k); keys.push(k); }
    return { keys, rows: flat, count: flat.length };
  } catch (e) {
    return { error: e.message };
  }
});

ipcMain.handle('match-xml', async (_evt, dayData, xmlRecords) => {
  const targetDate = dayData.businessDate;
  const xmlIndex = {};
  let xmlOnDateCount = 0;
  const xmlDateSet = new Set();
  for (const r of xmlRecords) {
    if (r.BUSINESS_DATE) xmlDateSet.add(r.BUSINESS_DATE);
    if (targetDate && r.BUSINESS_DATE !== targetDate) continue;
    xmlOnDateCount++;
    const room = pad4(r.ROOM || '');
    const item = String(r.ARTICLE_DESCRIPTION || '').trim();
    const qty = parseInt(r.QUANTITY, 10) || 0;
    if (!room || !item) continue;
    if (!xmlIndex[room]) xmlIndex[room] = {};
    xmlIndex[room][item] = (xmlIndex[room][item] || 0) + qty;
  }

  const cells = [];
  const stats = { match: 0, qtyMismatch: 0, onlyExcel: 0, onlyXml: 0 };
  const visited = new Set();

  for (const room of dayData.rooms) {
    for (const it of dayData.items) {
      const itemName = it.name;
      const excelQty = (dayData.grid[room] && dayData.grid[room][itemName]) || 0;
      const xmlQty = (xmlIndex[room] && xmlIndex[room][itemName]) || 0;
      visited.add(`${room}|${itemName}`);
      let status;
      if (excelQty === 0 && xmlQty === 0) status = 'empty';
      else if (excelQty === xmlQty) { status = 'match'; stats.match++; }
      else if (xmlQty === 0) { status = 'only-excel'; stats.onlyExcel++; }
      else if (excelQty === 0) { status = 'only-xml'; stats.onlyXml++; }
      else { status = 'qty-mismatch'; stats.qtyMismatch++; }
      cells.push({ room, item: itemName, excelQty, xmlQty, status });
    }
  }

  const extras = [];
  for (const room of Object.keys(xmlIndex)) {
    for (const item of Object.keys(xmlIndex[room])) {
      if (!visited.has(`${room}|${item}`)) {
        extras.push({ room, item, xmlQty: xmlIndex[room][item] });
        stats.onlyXml++;
      }
    }
  }

  return {
    cells, extras, stats, targetDate, xmlOnDateCount,
    availableXmlDates: [...xmlDateSet].sort(),
  };
});
