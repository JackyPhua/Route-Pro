const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('routePlannerElectron', {
  load: () => ipcRenderer.invoke('route-planner:load'),
  save: (json) => ipcRenderer.invoke('route-planner:save', json),
  saveSync: (json) => {
    ipcRenderer.sendSync('route-planner:save-sync', json);
  },
  /** 必須 sendSync：send 為異步時主進程尚未 focus 完頁面已對 input.focus() → 鍵盤仍不收字（須 minimize 才恢復） */
  focusWindow: () => {
    ipcRenderer.sendSync('route-planner:focus-window');
  },
});
