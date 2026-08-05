const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  checkEngines: () => ipcRenderer.invoke('check-engines'),
  selectSofficeManually: () => ipcRenderer.invoke('select-soffice-manually'),
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  selectOutputFolder: () => ipcRenderer.invoke('select-output-folder'),
  openFolder: (folderPath) => ipcRenderer.invoke('open-folder', folderPath),
  convertFiles: (payload) => ipcRenderer.invoke('convert-files', payload),
  onProgress: (callback) => ipcRenderer.on('conversion-progress', (_event, data) => callback(data))
});
