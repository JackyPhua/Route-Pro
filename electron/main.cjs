const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

const DATA_FILENAME = 'route-planner-state.json';

/** 部分電腦獨顯／驅動會間歇性空白渲染；使用者可在捷徑或系統環境變數設 ROUTE_PLANNER_DISABLE_GPU=1 */
if (process.env.ROUTE_PLANNER_DISABLE_GPU === '1') {
  app.disableHardwareAcceleration();
}

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

/** 避免頁面載入失敗／渲染崩潰時只見白屏、無任何提示 */
function attachRendererDiagnostics(win) {
  win.webContents.on(
    'did-fail-load',
    (_event, errorCode, errorDescription, validatedURL, isMainFrame) => {
      // -3 ERR_ABORTED：取消導航等，不算故障
      if (!isMainFrame || errorCode === -3) return;
      dialog.showErrorBox(
        'Route Planner — failed to load',
        [
          'The app UI did not load.',
          '',
          `Error ${errorCode}: ${errorDescription}`,
          `URL: ${validatedURL}`,
          '',
          'Try reinstalling, or exclude Route Planner from antivirus. Check that dist files were packaged.',
        ].join('\n'),
      );
    },
  );

  win.webContents.on('render-process-gone', (_event, details) => {
    if (details.reason === 'clean-exit') return;
    dialog.showMessageBoxSync(win, {
      type: 'error',
      title: 'Route Planner',
      message: 'The window stopped responding.',
      detail: [
        `Reason: ${details.reason}${details.exitCode != null ? ` (exit ${details.exitCode})` : ''}`,
        '',
        'Restart the app. On some PCs set environment variable ROUTE_PLANNER_DISABLE_GPU=1 before launch.',
      ].join('\n'),
      buttons: ['OK'],
    });
  });
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

  attachRendererDiagnostics(win);

  win.once('ready-to-show', () => win.show());

  const html = path.join(__dirname, '..', 'dist', 'index.html');
  win.loadFile(html).catch((err) => {
    dialog.showErrorBox(
      'Route Planner — missing files',
      `Could not open:\n${html}\n\n${err?.message ?? String(err)}\n\nRebuild with npm run electron:dist or reinstall.`,
    );
  });
}

process.on('uncaughtException', (err) => {
  console.error('route-planner: main uncaughtException', err);
});

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
