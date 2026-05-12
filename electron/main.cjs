const { app, BrowserWindow, ipcMain } = require('electron');
const fs = require('fs');
const path = require('path');

const DATA_FILENAME = 'route-planner-state.json';

function storePath() {
  return path.join(app.getPath('userData'), DATA_FILENAME);
}

function readStore() {
  const p = storePath();
  try {
    if (!fs.existsSync(p)) return null;
    const raw = fs.readFileSync(p, 'utf8');
    return raw && raw.trim().length > 0 ? raw : null;
  } catch (err) {
    console.warn('route-planner: read store failed', err);
    return null;
  }
}

function writeStore(raw) {
  const p = storePath();
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, String(raw ?? ''), 'utf8');
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'Route Planner',
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.cjs'),
    },
    show: false,
  });

  win.once('ready-to-show', () => win.show());

  const html = path.join(__dirname, '..', 'dist', 'index.html');
  win.loadFile(html);
}

app.whenReady().then(() => {
  ipcMain.handle('route-planner:load', () => readStore());
  ipcMain.handle('route-planner:save', (_e, json) => {
    try {
      writeStore(json);
    } catch (err) {
      console.error('route-planner: async save failed', err);
      throw err;
    }
  });
  ipcMain.on('route-planner:save-sync', (_e, json) => {
    try {
      writeStore(json);
    } catch (err) {
      console.error('route-planner: sync save failed', err);
    }
  });

  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
