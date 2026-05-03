const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { XMLParser } = require('fast-xml-parser');

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 780,
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

ipcMain.handle('get-rooms', async (_evt, filePath, day) => {
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
    const rooms = [];
    for (const name of matching) {
      const ws = wb.Sheets[name];
      if (!ws['!ref']) continue;
      const ref = XLSX.utils.decode_range(ws['!ref']);
      // Row 5 in spreadsheet = r:4; columns H onward = c:7..end
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
    return { rooms, csv: rooms.join(','), sheets: matching };
  } catch (e) {
    return { error: e.message };
  }
});

// Recursively walk and collect all G_TRX_CHAR_DATE objects
function collectRecords(node, out) {
  if (node === null || node === undefined) return;
  if (Array.isArray(node)) {
    for (const item of node) collectRecords(item, out);
    return;
  }
  if (typeof node !== 'object') return;
  for (const [k, v] of Object.entries(node)) {
    if (k === 'G_TRX_CHAR_DATE') {
      if (Array.isArray(v)) out.push(...v);
      else out.push(v);
    } else {
      collectRecords(v, out);
    }
  }
}

ipcMain.handle('parse-xml', async (_evt, filePath) => {
  try {
    const xml = fs.readFileSync(filePath, 'utf8');
    const parser = new XMLParser({
      ignoreAttributes: true,
      trimValues: true,
      parseTagValue: false, // keep all values as strings for predictable display
    });
    const obj = parser.parse(xml);
    const records = [];
    collectRecords(obj, records);
    // Normalize: flatten each record to plain {key: string} map
    const flat = records.map((r) => {
      const out = {};
      for (const [k, v] of Object.entries(r)) {
        if (v === null || v === undefined) out[k] = '';
        else if (typeof v === 'object') out[k] = JSON.stringify(v);
        else out[k] = String(v).trim();
      }
      return out;
    });
    // Compute union of keys preserving first-seen order
    const keys = [];
    const seen = new Set();
    for (const r of flat) {
      for (const k of Object.keys(r)) {
        if (!seen.has(k)) {
          seen.add(k);
          keys.push(k);
        }
      }
    }
    return { keys, rows: flat, count: flat.length };
  } catch (e) {
    return { error: e.message };
  }
});
