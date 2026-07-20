/** 將上次選取的試算簿二進位/CSV 存 IndexedDB ，重開後可不重新選檔即可換工作表 */

const DB_NAME = 'route-planner-import-sheet-cache-v1';
const STORE = 'store';
const STORE_KEY_LAST = 'last';
const SCHEMA_VERSION = 1;

/** @typedef {{ version:number, fileName:string, sheetIndex:number, kind:'csv'|'binary', payload:string|ArrayBuffer }} CachedImportRecord */

function openDb() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') {
      reject(new Error('indexedDB unavailable'));
      return;
    }
    const req = indexedDB.open(DB_NAME, 1);
    req.onerror = () => reject(req.error);
    req.onsuccess = () => resolve(req.result);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };
  });
}

/**
 * @param {CachedImportRecord} record
 */
export async function saveCachedSpreadsheet(record) {
  try {
    const db = await openDb();
    const rec = { ...record, version: SCHEMA_VERSION };
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
      tx.objectStore(STORE).put(rec, STORE_KEY_LAST);
    });
    db.close();
  } catch (err) {
    console.warn('route-planner: saveCachedSpreadsheet failed', err);
  }
}

/** @returns {Promise<CachedImportRecord|null>} */
export async function loadCachedSpreadsheet() {
  try {
    if (typeof indexedDB === 'undefined') return null;
    const db = await openDb();
    const rec = await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      tx.onerror = () => reject(tx.error);
      const rq = tx.objectStore(STORE).get(STORE_KEY_LAST);
      rq.onsuccess = () => resolve(rq.result ?? null);
      rq.onerror = () => reject(rq.error);
    });
    db.close();
    return rec &&
      typeof rec === 'object' &&
      typeof rec.kind === 'string' &&
      (rec.kind === 'csv' ? typeof rec.payload === 'string' : rec.payload instanceof ArrayBuffer)
      ? /** @type {CachedImportRecord} */ (rec)
      : null;
  } catch (err) {
    console.warn('route-planner: loadCachedSpreadsheet failed', err);
    return null;
  }
}

export async function clearCachedSpreadsheet() {
  try {
    if (typeof indexedDB === 'undefined') return;
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
      tx.objectStore(STORE).delete(STORE_KEY_LAST);
    });
    db.close();
  } catch (err) {
    console.warn('route-planner: clearCachedSpreadsheet failed', err);
  }
}

/**
 * @param {number} sheetIndex
 */
export async function patchCachedSpreadsheetSheetIndex(sheetIndex) {
  const cur = await loadCachedSpreadsheet();
  if (!cur) return;
  const idx = Math.max(0, Math.floor(Number(sheetIndex)) || 0);
  await saveCachedSpreadsheet({ ...cur, sheetIndex: idx });
}
