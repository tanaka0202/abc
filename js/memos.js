// メモ（REQ-4 / B-5 / B-6）。端末内に保存し、会議フォルダがあればOneDriveにも保存する。
import { kvGet, kvSet } from './db.js';
import { getSettings } from './settings.js';
import { uuid, debounce } from './util.js';
import { writeMyMemos } from './sync.js';
import { mergeMemos } from './textproc.js';

const key = (sessionId) => `memos:${sessionId}`;
const uploaders = new Map();
const locks = new Map();

/** 同じ会議のメモの読み書きを1件ずつ順番に行う（連続入力で書き込みが消えないように） */
function serial(sessionId, fn) {
  const prev = locks.get(sessionId) || Promise.resolve();
  const next = prev.then(fn, fn);
  locks.set(sessionId, next.catch(() => {}));
  return next;
}

export async function getMyMemos(sessionId) { return (await kvGet(key(sessionId), [])) || []; }

async function setMyMemos(sessionId, list, folderPath) {
  await kvSet(key(sessionId), list);
  if (folderPath && getSettings().onedrive.clientId) scheduleUpload(sessionId, folderPath);
}

function scheduleUpload(sessionId, folderPath) {
  let u = uploaders.get(sessionId);
  if (!u) {
    u = debounce(async (fp) => {
      try { await writeMyMemos(fp, await getMyMemos(sessionId)); } catch (e) { console.warn('メモの保存（OneDrive）に失敗', e); }
    }, 1500);
    uploaders.set(sessionId, u);
  }
  u(folderPath);
}

/** 未送信のメモをすぐ送る */
export async function flushMemos(sessionId, folderPath) {
  if (!folderPath) return;
  const u = uploaders.get(sessionId);
  if (u) await u.flush(folderPath);
  else await writeMyMemos(folderPath, await getMyMemos(sessionId));
}

/**
 * メモを追加する。
 * t0: 書き始めた時刻（ミリ秒）。startedAt と合わせて録音内の位置を決める。
 * offset を直接渡した場合（会議後に再生位置へ付けるメモ）はそれを使う。
 */
export function addMemo(target, opts) { return serial(target.id, () => addMemoNow(target, opts)); }

async function addMemoNow(target, { text, kind = 'memo', t0 = Date.now(), offset = null }) {
  const st = getSettings();
  const list = await getMyMemos(target.id);
  let off = offset;
  if (off == null && Number.isFinite(target.startedAt)) off = Math.round((t0 - target.startedAt) / 100) / 10;
  const m = {
    id: `${st.deviceId}-${uuid().slice(0, 8)}`, sessionId: target.id, deviceId: st.deviceId, deviceName: st.deviceName,
    t: t0, offset: off, text: String(text || '').trim() || (kind === 'star' ? '★要確認' : ''), kind,
    createdAt: Date.now(), updatedAt: Date.now(),
  };
  list.push(m);
  await setMyMemos(target.id, list, target.folderPath);
  return m;
}

export function updateMemo(target, id, patch) { return serial(target.id, () => updateMemoNow(target, id, patch)); }

async function updateMemoNow(target, id, patch) {
  const list = await getMyMemos(target.id);
  const m = list.find((x) => x.id === id);
  if (!m) return null;
  Object.assign(m, patch, { updatedAt: Date.now() });
  await setMyMemos(target.id, list, target.folderPath);
  return m;
}

export async function deleteMemo(target, id) { return updateMemo(target, id, { deleted: true }); }

/** 録音開始前に書いたメモに、開始時刻が決まった時点で位置を付ける */
export function fixOffsets(target) { return serial(target.id, () => fixOffsetsNow(target)); }

async function fixOffsetsNow(target) {
  const list = await getMyMemos(target.id);
  let changed = false;
  for (const m of list) {
    if (m.offset == null && Number.isFinite(target.startedAt) && Number.isFinite(m.t)) {
      m.offset = Math.round((m.t - target.startedAt) / 100) / 10;
      changed = true;
    }
  }
  if (changed) await setMyMemos(target.id, list, target.folderPath);
}

/** 自分のメモと、OneDrive上の他端末のメモ（端末ごとの配列の配列）を統合 */
export function combine(mine, others = []) {
  return mergeMemos(mine, ...others);
}
