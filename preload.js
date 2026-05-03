const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectFile: (kind) => ipcRenderer.invoke('select-file', kind),
  getDayData: (filePath, day) => ipcRenderer.invoke('get-day-data', filePath, day),
  parseXml: (filePath) => ipcRenderer.invoke('parse-xml', filePath),
  matchXml: (dayData, xmlRecords) => ipcRenderer.invoke('match-xml', dayData, xmlRecords),
});
