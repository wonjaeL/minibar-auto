const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectFile: (kind) => ipcRenderer.invoke('select-file', kind),
  getRooms: (filePath, day) => ipcRenderer.invoke('get-rooms', filePath, day),
  parseXml: (filePath) => ipcRenderer.invoke('parse-xml', filePath),
});
