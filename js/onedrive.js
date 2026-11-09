// OneDrive（Microsoft Graph）。個人用OneDriveの指定フォルダに読み書きする。
import { msGetToken } from './auth.js';
import { getSettings } from './settings.js';
import { AppError, sleep } from './util.js';

const GRAPH = 'https://graph.microsoft.com/v1.0/me/drive';
const SMALL_LIMIT = 4 * 1024 * 1024;
const CHUNK = 320 * 1024 * 16; // 5MiB（320KiBの倍数である必要がある）

const encPath = (p) => String(p).split('/').filter(Boolean).map(encodeURIComponent).join('/');
export const joinPath = (...parts) => parts.flatMap((p) => String(p || '').split('/')).filter(Boolean).join('/');

async function token() { return msGetToken(getSettings().onedrive.clientId); }

async function graph(method, url, { body, headers = {}, raw = false, allow404 = false } = {}) {
  for (let attempt = 0; attempt < 5; attempt++) {
    const t = await token();
    let r;
    try {
      r = await fetch(url.startsWith('http') ? url : `${GRAPH}${url}`, {
        method,
        headers: { Authorization: `Bearer ${t}`, ...headers },
        body,
      });
    } catch (e) {
      if (attempt < 4) { await sleep(1000 * 2 ** attempt); continue; }
      throw new AppError('OneDriveに接続できませんでした（ネットワークを確認してください）。', { cause: e });
    }
    if (r.status === 404 && allow404) return null;
    if (r.status === 429 || r.status === 503 || r.status === 502 || r.status === 504) {
      const ra = Number(r.headers.get('Retry-After')) || 2 ** attempt;
      await sleep(Math.min(30, ra) * 1000);
      continue;
    }
    if (r.status === 401 && attempt === 0) continue; // トークン更新の行き違い
    if (!r.ok) {
      const j = await r.json().catch(() => ({}));
      const msg = (j.error && (j.error.message || j.error.code)) || r.status;
      const err = new AppError(`OneDriveの操作に失敗しました（${msg}）`);
      err.status = r.status;
      throw err;
    }
    if (raw) return r;
    if (r.status === 204) return null;
    return r.json();
  }
  throw new AppError('OneDriveが混み合っています。しばらくしてから再試行してください。');
}

export async function getItem(path) {
  return graph('GET', `/root:/${encPath(path)}`, { allow404: true });
}

export async function getMe() {
  const t = await token();
  const r = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { Authorization: `Bearer ${t}` } });
  if (!r.ok) throw new AppError('OneDriveのアカウント情報を取得できませんでした。');
  return r.json();
}

/** フォルダがなければ親から順に作る */
export async function ensureFolder(path) {
  const parts = String(path).split('/').filter(Boolean);
  let cur = '';
  for (const name of parts) {
    const next = cur ? `${cur}/${name}` : name;
    const item = await getItem(next);
    if (!item) {
      const parentUrl = cur ? `/root:/${encPath(cur)}:/children` : '/root/children';
      try {
        await graph('POST', parentUrl, {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
        });
      } catch (e) {
        if (e.status !== 409) throw e; // 同時に作られた場合は問題なし
      }
    } else if (!item.folder) {
      throw new AppError(`「${next}」はフォルダではありません。`);
    }
    cur = next;
  }
  return getItem(cur);
}

/** 同名フォルダがあれば _2, _3 … を付けて新しいフォルダを作る */
export async function createUniqueFolder(parent, name) {
  await ensureFolder(parent);
  for (let i = 1; i < 50; i++) {
    const n = i === 1 ? name : `${name}_${i}`;
    const exists = await getItem(joinPath(parent, n));
    if (!exists) {
      try {
        await graph('POST', `/root:/${encPath(parent)}:/children`, {
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: n, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }),
        });
        return { name: n, path: joinPath(parent, n) };
      } catch (e) {
        if (e.status !== 409) throw e;
      }
    }
  }
  throw new AppError('保存先フォルダを作れませんでした。');
}

/** 文字列やBlobを保存（4MB以上は分割アップロード） */
export async function upload(path, data, contentType = 'application/octet-stream', onProgress) {
  const blob = data instanceof Blob ? data : new Blob([data], { type: contentType });
  if (blob.size < SMALL_LIMIT) {
    return graph('PUT', `/root:/${encPath(path)}:/content`, { headers: { 'Content-Type': contentType }, body: blob });
  }
  return uploadLarge(path, blob, onProgress);
}

export const uploadText = (path, text, type = 'text/plain;charset=utf-8') => upload(path, text, type);
export const uploadJson = (path, obj) => upload(path, JSON.stringify(obj, null, 1), 'application/json');

async function uploadLarge(path, blob, onProgress) {
  const sess = await graph('POST', `/root:/${encPath(path)}:/createUploadSession`, {
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
  });
  const url = sess.uploadUrl;
  let offset = 0;
  let result = null;
  while (offset < blob.size) {
    const end = Math.min(offset + CHUNK, blob.size);
    let ok = false;
    for (let attempt = 0; attempt < 6 && !ok; attempt++) {
      try {
        // 分割アップロード先のURLは認証済みなので、Authorizationヘッダは付けない
        const r = await fetch(url, {
          method: 'PUT',
          headers: { 'Content-Range': `bytes ${offset}-${end - 1}/${blob.size}` },
          body: blob.slice(offset, end),
        });
        if (r.status === 200 || r.status === 201) { result = await r.json(); ok = true; break; }
        if (r.status === 202) {
          const j = await r.json().catch(() => ({}));
          const next = j.nextExpectedRanges && j.nextExpectedRanges[0];
          if (next) { const n = Number(String(next).split('-')[0]); if (Number.isFinite(n) && n > offset) { offset = n; ok = true; continue; } }
          offset = end; ok = true; continue;
        }
        if (r.status === 416) { // 既に受け取り済みの範囲
          const st = await fetch(url).then((x) => x.json()).catch(() => null);
          const next = st && st.nextExpectedRanges && st.nextExpectedRanges[0];
          if (next) { offset = Number(String(next).split('-')[0]); ok = true; continue; }
        }
        if (r.status === 404) throw new AppError('アップロードの有効期限が切れました。再試行してください。');
      } catch (e) {
        if (e instanceof AppError) throw e;
      }
      await sleep(1000 * 2 ** attempt);
    }
    if (!ok) throw new AppError('音声のアップロードに失敗しました（ネットワークを確認して再試行してください）。');
    if (result) break;
    if (onProgress) onProgress(offset / blob.size);
  }
  if (onProgress) onProgress(1);
  return result;
}

/** ファイルの内容を取得（Graphの /content は転送のためブラウザでは使えないので downloadUrl を使う） */
export async function download(path, as = 'text') {
  const item = await getItem(path);
  if (!item || item.folder) return null;
  const url = item['@microsoft.graph.downloadUrl'];
  if (!url) return null;
  const r = await fetch(url);
  if (!r.ok) throw new AppError(`OneDriveから読み込めませんでした（${r.status}）`);
  const body = as === 'blob' ? await r.blob() : as === 'json' ? await r.json() : await r.text();
  return { body, eTag: item.eTag, item };
}

export async function readJson(path) {
  try {
    const res = await download(path, 'json');
    return res ? res.body : null;
  } catch (e) {
    if (e instanceof SyntaxError) return null;
    throw e;
  }
}

export async function readText(path) {
  const res = await download(path, 'text');
  return res ? res.body : null;
}

export async function listChildren(path) {
  const out = [];
  let url = `/root:/${encPath(path)}:/children?$top=200&$select=id,name,folder,file,size,webUrl,lastModifiedDateTime,eTag`;
  while (url) {
    const j = await graph('GET', url, { allow404: true });
    if (!j) break;
    out.push(...(j.value || []));
    url = j['@odata.nextLink'] || null;
  }
  return out;
}

export async function search(path, q) {
  const safe = String(q).replace(/'/g, "''");
  const j = await graph('GET', `/root:/${encPath(path)}:/search(q='${encodeURIComponent(safe)}')?$top=100`, { allow404: true });
  return (j && j.value) || [];
}

/** 一時的な再生用URL（約1時間有効） */
export async function downloadUrl(path) {
  const item = await getItem(path);
  return item ? item['@microsoft.graph.downloadUrl'] || null : null;
}

export async function webUrl(path) {
  const item = await getItem(path);
  return item ? item.webUrl : null;
}
