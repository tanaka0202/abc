// 端末内保存（IndexedDB）。録音の一時データ・会議・設定・単語帳・利用履歴を保持する。

const DB_NAME = 'mtgrec';
const DB_VERSION = 1;
let dbPromise = null;

function open() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
      if (!db.objectStoreNames.contains('sessions')) db.createObjectStore('sessions', { keyPath: 'id' });
      if (!db.objectStoreNames.contains('chunks')) {
        const s = db.createObjectStore('chunks', { keyPath: ['sessionId', 'seq'] });
        s.createIndex('bySession', 'sessionId');
      }
      if (!db.objectStoreNames.contains('files')) {
        const f = db.createObjectStore('files', { keyPath: 'id' });
        f.createIndex('bySession', 'sessionId');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

function tx(store, mode, fn) {
  return open().then((db) => new Promise((resolve, reject) => {
    const t = db.transaction(store, mode);
    const s = t.objectStore(store);
    let result;
    Promise.resolve(fn(s)).then((r) => { result = r; }).catch(reject);
    t.oncomplete = () => resolve(result);
    t.onerror = () => reject(t.error);
    t.onabort = () => reject(t.error || new Error('保存が中断されました'));
  }));
}

const wrap = (req) => new Promise((resolve, reject) => {
  req.onsuccess = () => resolve(req.result);
  req.onerror = () => reject(req.error);
});

// ---- kv ----
export async function kvGet(key, fallback = undefined) {
  const v = await tx('kv', 'readonly', (s) => wrap(s.get(key)));
  return v === undefined ? fallback : v;
}
export function kvSet(key, value) { return tx('kv', 'readwrite', (s) => wrap(s.put(value, key))); }
export function kvDel(key) { return tx('kv', 'readwrite', (s) => wrap(s.delete(key))); }

// ---- sessions ----
export function putSession(session) {
  session.updatedAt = Date.now();
  return tx('sessions', 'readwrite', (s) => wrap(s.put(session)));
}
export function getSession(id) { return tx('sessions', 'readonly', (s) => wrap(s.get(id))); }
export async function listSessions() {
  const all = await tx('sessions', 'readonly', (s) => wrap(s.getAll()));
  return (all || []).sort((a, b) => (b.startedAt || 0) - (a.startedAt || 0));
}
export function deleteSessionRecord(id) { return tx('sessions', 'readwrite', (s) => wrap(s.delete(id))); }

// ---- chunks（録音の一時データ） ----
export function addChunk(sessionId, seq, buffer) {
  return tx('chunks', 'readwrite', (s) => wrap(s.put({ sessionId, seq, data: buffer })));
}

/** セッションの録音データを順番どおりに取り出す（ArrayBuffer の配列） */
export function getChunks(sessionId) {
  return tx('chunks', 'readonly', (s) => new Promise((resolve, reject) => {
    const out = [];
    const range = IDBKeyRange.bound([sessionId, -Infinity], [sessionId, Infinity]);
    const req = s.openCursor(range);
    req.onsuccess = () => {
      const c = req.result;
      if (c) { out.push(c.value.data); c.continue(); } else resolve(out);
    };
    req.onerror = () => reject(req.error);
  }));
}

export function deleteChunks(sessionId) {
  return tx('chunks', 'readwrite', (s) => {
    const range = IDBKeyRange.bound([sessionId, -Infinity], [sessionId, Infinity]);
    return wrap(s.delete(range));
  });
}

// ---- files（会議資料など） ----
export function putFile(rec) { return tx('files', 'readwrite', (s) => wrap(s.put(rec))); }
export function getFiles(sessionId) {
  return tx('files', 'readonly', (s) => wrap(s.index('bySession').getAll(sessionId)));
}
export function deleteFile(id) { return tx('files', 'readwrite', (s) => wrap(s.delete(id))); }

/** 端末内のデータをすべて消去 */
export async function wipeAll() {
  const db = await open();
  db.close();
  dbPromise = null;
  await new Promise((resolve, reject) => {
    const req = indexedDB.deleteDatabase(DB_NAME);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
    req.onblocked = () => resolve();
  });
}

/** ストレージを消されにくくする（対応ブラウザのみ） */
export async function requestPersistence() {
  try {
    if (navigator.storage && navigator.storage.persist) return await navigator.storage.persist();
  } catch { /* 非対応 */ }
  return false;
}
