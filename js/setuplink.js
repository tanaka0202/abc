// ほかの端末への設定の受け渡し（QRコードに入れる内容）。純粋な関数だけを置く。
// 秘密にすべきものはGeminiのAPIキーだけ。アプリID・クライアントIDは、ブラウザに渡る公開の識別子。

export const SETUP_PREFIX = '#setup=';
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DEFAULT_ROOT = '会議録音アーカイブ';

function b64urlEncode(str) {
  const bytes = new TextEncoder().encode(str);
  let bin = '';
  bytes.forEach((b) => { bin += String.fromCharCode(b); });
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function b64urlDecode(s) {
  const bin = atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4));
  return new TextDecoder().decode(Uint8Array.from(bin, (c) => c.charCodeAt(0)));
}

/** 設定から、渡す内容を作る（既定値は省いて短くする） */
export function setupPayload(st, includeKey) {
  const p = { v: 1, od: st.onedrive.clientId };
  if (st.onedrive.rootPath && st.onedrive.rootPath !== DEFAULT_ROOT) p.r = st.onedrive.rootPath;
  if (includeKey && st.gemini.apiKey) { p.g = st.gemini.apiKey; if (st.gemini.paidConfirmed) p.p = 1; }
  if (st.google.clientId) p.gc = st.google.clientId;
  if (st.google.account) p.ga = st.google.account;
  return p;
}

export const encodeSetup = (p) => b64urlEncode(JSON.stringify(p));

/** 受け取った内容を読み、形式が正しいものだけを返す（不正なものは null） */
export function parseSetup(raw) {
  let p;
  try { p = JSON.parse(b64urlDecode(String(raw || ''))); } catch { return null; }
  if (!p || p.v !== 1 || typeof p.od !== 'string' || !GUID_RE.test(p.od)) return null;
  const str = (v, re, max) => (typeof v === 'string' && v.length <= max && re.test(v) ? v : null);
  const out = { od: p.od };
  if (p.r !== undefined) { out.r = str(p.r, /^[^\\:*?"<>|]+$/, 200); if (!out.r) return null; }
  if (p.g !== undefined) { out.g = str(p.g, /^[A-Za-z0-9._-]+$/, 400); if (!out.g) return null; out.p = p.p === 1; }
  if (p.gc !== undefined) { out.gc = str(p.gc, /^[0-9a-z-]+\.apps\.googleusercontent\.com$/, 200); if (!out.gc) return null; }
  if (p.ga !== undefined) { out.ga = str(p.ga, /^[^\s@]+@[^\s@]+$/, 200); if (!out.ga) return null; }
  return out;
}
