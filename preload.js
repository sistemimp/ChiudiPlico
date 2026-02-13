const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectFile: () => ipcRenderer.invoke('select-file'),
  processFile: (data) => ipcRenderer.invoke('process-file', data),
  getPathForFile: (file) => webUtils.getPathForFile(file),
});
