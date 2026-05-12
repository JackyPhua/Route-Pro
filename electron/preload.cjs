const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('routePlannerElectron', {
  load: () => ipcRenderer.invoke('route-planner:load'),
  save: (json) => ipcRenderer.invoke('route-planner:save', json),
  saveSync: (json) => {
    ipcRenderer.sendSync('route-planner:save-sync', json);
  },
});
