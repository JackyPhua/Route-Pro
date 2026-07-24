const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

/** GitHub Releases auto-update (NSIS installed build only; skipped in dev / portable). */
function setupAutoUpdate(win) {
  if (!app.isPackaged || process.platform !== 'win32') return;
  try {
    const { autoUpdater } = require('electron-updater');
    autoUpdater.autoDownload = true;
    autoUpdater.autoInstallOnAppQuit = true;

    autoUpdater.on('error', (err) => {
      console.warn('route-planner: autoUpdater error', err?.message ?? err);
    });

    autoUpdater.on('update-available', (info) => {
      console.info('route-planner: update available', info?.version ?? info);
    });

    autoUpdater.on('update-downloaded', (info) => {
      const parent =
        win && !win.isDestroyed() ? win : BrowserWindow.getFocusedWindow() ?? undefined;
      dialog
        .showMessageBox(parent, {
          type: 'info',
          title: 'Route Planner update',
          message: `Version ${info.version} is ready to install.`,
          detail: 'Restart now to apply the update, or choose Later to install on next exit.',
          buttons: ['Restart now', 'Later'],
          defaultId: 0,
          cancelId: 1,
        })
        .then(({ response }) => {
          if (response === 0) autoUpdater.quitAndInstall(false, true);
        })
        .catch(() => {});
    });

    autoUpdater.checkForUpdatesAndNotify().catch((err) => {
      console.warn('route-planner: checkForUpdates failed', err?.message ?? err);
    });
  } catch (err) {
    console.warn('route-planner: autoUpdater unavailable', err?.message ?? err);
  }
}

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

/** @type {import('electron').BrowserWindow | null} */
let mainWindow = null;

function resolveAppIcon() {
  const ico = path.join(__dirname, 'icon.ico');
  const png = path.join(__dirname, 'icon.png');
  if (fs.existsSync(ico)) return ico;
  if (fs.existsSync(png)) return png;
  return undefined;
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'Route Planner',
    icon: resolveAppIcon(),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      preload: path.join(__dirname, 'preload.cjs'),
    },
    show: false,
  });

  mainWindow = win;
  win.on('closed', () => {
    if (mainWindow === win) mainWindow = null;
  });

  attachRendererDiagnostics(win);

  win.once('ready-to-show', () => {
    win.maximize();
    win.show();
  });

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
  ipcMain.on('route-planner:save-sync', (event, json) => {
    try {
      writeStore(json);
    } catch (err) {
      console.error('route-planner: sync save failed', err);
    }
    /** sendSync 須設定 returnValue ，否則渲染進程可能一直阻塞 → 對話框不关、視窗似「當機」 */
    event.returnValue = null;
  });

  /** 原生 alert / 覆蓋層後部分環境鍵盤焦點留在 OS ，需把焦點拉回視窗與 webContents（等同使用者 minimize 再開） */
  ipcMain.on('route-planner:focus-window', (event) => {
    try {
      const w =
        mainWindow && !mainWindow.isDestroyed()
          ? mainWindow
          : BrowserWindow.getAllWindows()[0];
      if (!w || w.isDestroyed()) return;
      if (w.isMinimized()) w.restore();
      w.show();
      w.focus();
      w.webContents.focus();
    } catch (err) {
      console.warn('route-planner: focus-window failed', err);
    }
    /** sendSync 須設定 returnValue ，否則渲染進程會阻塞 */
    event.returnValue = null;
  });

  createWindow();
  if (mainWindow) setupAutoUpdate(mainWindow);
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
