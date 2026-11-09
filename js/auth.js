// サインイン（Microsoft：PKCE方式、Google：ブラウザ向けトークン方式）。
// パスワードはアプリに渡らない。トークンは端末内（IndexedDB）にだけ保存する。
import { kvGet, kvSet, kvDel } from './db.js';
import { AppError } from './util.js';

const CHANNEL = 'mtgrec-auth';
const RESULT_KEY = 'mtgrec-auth-result';
const PENDING_KEY = 'mtgrec-auth-pending';

const MS_AUTHORITY = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const MS_SCOPES = 'Files.ReadWrite User.Read offline_access openid profile';
const G_AUTH = 'https://accounts.google.com/o/oauth2/v2/auth';
export const G_SCOPES = 'https://www.googleapis.com/auth/calendar.events';

export class NeedLogin extends AppError {
  constructor(provider, message) {
    super(message || (provider === 'ms' ? 'OneDriveへのサインインが必要です。' : 'Googleカレンダーへのサインインが必要です。'), { code: `need_login_${provider}` });
    this.provider = provider;
  }
}

export function redirectUri() {
  return new URL('./auth.html', document.baseURI).href.split('#')[0];
}

function randomString(len = 48) {
  const a = new Uint8Array(len);
  crypto.getRandomValues(a);
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  return Array.from(a, (x) => chars[x % chars.length]).join('');
}

async function pkceChallenge(verifier) {
  const data = new TextEncoder().encode(verifier);
  const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', data));
  let s = '';
  for (const b of hash) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

/** 利用者の操作の直後に呼ぶこと（ポップアップのブロックを避けるため、先に空の窓を開く） */
function openPopup() {
  const w = 480;
  const h = 680;
  const left = Math.max(0, (window.screenX || 0) + ((window.outerWidth || w) - w) / 2);
  const top = Math.max(0, (window.screenY || 0) + ((window.outerHeight || h) - h) / 2);
  try {
    return window.open('', 'mtgrec-auth', `width=${w},height=${h},left=${left},top=${top}`);
  } catch {
    return null;
  }
}

/**
 * サインインの結果を待つ。watch に開いた窓を渡すと、結果が届かないまま窓が閉じられたときに、
 * 5分待たせずに知らせる（サインイン画面にエラーが出て、利用者が閉じた場合など）。
 */
function waitForResult(state, timeoutMs = 5 * 60 * 1000, watch = null) {
  return new Promise((resolve, reject) => {
    let done = false;
    let bc = null;
    const finish = (data) => {
      if (done || !data || data.state !== state) return;
      done = true;
      cleanup();
      try { localStorage.removeItem(RESULT_KEY); } catch { /* 無視 */ }
      resolve(data);
    };
    const onMsg = (e) => { if (e.data && e.data.type === CHANNEL) finish(e.data.data); };
    const onStorage = (e) => {
      if (e.key === RESULT_KEY && e.newValue) { try { finish(JSON.parse(e.newValue).data); } catch { /* 無視 */ } }
    };
    const timer = setTimeout(() => { if (!done) { done = true; cleanup(); reject(new AppError('サインインがタイムアウトしました。もう一度お試しください。')); } }, timeoutMs);
    let closedAt = 0;
    const poll = setInterval(() => {
      try { const raw = localStorage.getItem(RESULT_KEY); if (raw) finish(JSON.parse(raw).data); } catch { /* 無視 */ }
      let closed = false;
      try { closed = !!(watch && watch.closed); } catch { closed = false; }
      if (!closed || done) return;
      if (!closedAt) { closedAt = Date.now(); return; }
      // 戻り先ページは結果を渡してから自分で閉じるので、閉じてから少し待って結果がなければ失敗とする
      if (Date.now() - closedAt > 2000) {
        done = true;
        cleanup();
        const err = new AppError('サインインが完了しませんでした（サインイン画面が閉じられました）。画面に英語のエラーが出ていた場合は、手順書7章「サインインのエラー」で原因を確認してください。多くは、戻り先URLがMicrosoft Entraの登録と一致していないことが原因です。');
        err.code = 'popup_closed';
        reject(err);
      }
    }, 700);
    function cleanup() {
      clearTimeout(timer);
      clearInterval(poll);
      window.removeEventListener('message', onMsg);
      window.removeEventListener('storage', onStorage);
      if (bc) bc.close();
    }
    try { bc = new BroadcastChannel(CHANNEL); bc.onmessage = onMsg; } catch { bc = null; }
    window.addEventListener('message', onMsg);
    window.addEventListener('storage', onStorage);
  });
}

async function runFlow(popup, url, state, pending, { watchClosed = false } = {}) {
  if (!popup) {
    // ポップアップが使えない場合は、画面ごと移動して戻ってくる
    localStorage.setItem(PENDING_KEY, JSON.stringify({ ...pending, state, at: Date.now() }));
    location.assign(url);
    return new Promise(() => {}); // 戻ってくるまで待つ（画面が切り替わる）
  }
  popup.location.href = url;
  return waitForResult(state, undefined, watchClosed ? popup : null);
}

// ---------------------------------------------------------------- Microsoft
export async function msLogin(clientId) {
  if (!clientId) throw new AppError('OneDriveのアプリID（クライアントID）が設定されていません。');
  const popup = openPopup();
  const verifier = randomString(64);
  const state = randomString(24);
  const challenge = await pkceChallenge(verifier);
  const url = `${MS_AUTHORITY}/authorize?${new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'query',
    scope: MS_SCOPES,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    state,
    prompt: 'select_account',
  })}`;
  // Microsoftのサインイン画面は窓の開閉を確認できる（Googleの画面は、確認すると誤って「閉じた」と判定されることがあるため対象外）
  const data = await runFlow(popup, url, state, { provider: 'ms', verifier, clientId }, { watchClosed: true });
  return msFinish(clientId, data, verifier);
}

async function msFinish(clientId, data, verifier) {
  if (data.error) throw new AppError(`Microsoftのサインインに失敗しました：${data.error_description || data.error}`);
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: 'authorization_code',
    code: data.code,
    redirect_uri: redirectUri(),
    code_verifier: verifier,
    scope: MS_SCOPES,
  });
  const r = await fetch(`${MS_AUTHORITY}/token`, { method: 'POST', body });
  const j = await r.json().catch(() => ({}));
  if (!r.ok) throw new AppError(`Microsoftのサインインに失敗しました：${j.error_description || j.error || r.status}`);
  await saveMsTokens(j, true);
  return true;
}

async function saveMsTokens(j, fresh = false) {
  const prev = (await kvGet('ms_tokens')) || {};
  await kvSet('ms_tokens', {
    accessToken: j.access_token,
    refreshToken: j.refresh_token || prev.refreshToken,
    expiresAt: Date.now() + (Number(j.expires_in) || 3600) * 1000,
    refreshedAt: Date.now(),
    signedInAt: fresh ? Date.now() : (prev.signedInAt || Date.now()),
  });
}

let msRefreshing = null;
export async function msGetToken(clientId) {
  const t = await kvGet('ms_tokens');
  if (!t) throw new NeedLogin('ms');
  if (t.expiresAt - 60_000 > Date.now()) return t.accessToken;
  if (!t.refreshToken) throw new NeedLogin('ms');
  // ブラウザ向けアプリの更新用トークンは、最初のサインインから約24時間で使えなくなる。
  // 期限が近ければ通信せずにすぐ再サインインを求める（操作の直後に認証画面を開けるようにするため）
  if (t.signedInAt && Date.now() - t.signedInAt > 23.5 * 3600_000) throw new NeedLogin('ms', 'OneDriveのサインインの有効期限（約24時間）が切れました。再サインインしてください。');
  if (!msRefreshing) {
    msRefreshing = (async () => {
      const body = new URLSearchParams({ client_id: clientId, grant_type: 'refresh_token', refresh_token: t.refreshToken, scope: MS_SCOPES });
      const r = await fetch(`${MS_AUTHORITY}/token`, { method: 'POST', body });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) {
        if (j.error === 'invalid_grant' || j.error === 'interaction_required') {
          await kvDel('ms_tokens');
          throw new NeedLogin('ms', 'OneDriveのサインインの有効期限が切れました。再サインインしてください。');
        }
        throw new AppError(`OneDriveのサインインを更新できませんでした（${j.error || r.status}）`);
      }
      await saveMsTokens(j);
      return j.access_token;
    })().finally(() => { msRefreshing = null; });
  }
  return msRefreshing;
}

export async function msSignedIn() { return !!(await kvGet('ms_tokens')); }
export async function msLogout() { await kvDel('ms_tokens'); }

// ---------------------------------------------------------------- Google
export async function googleLogin(clientId, account = '') {
  if (!clientId) throw new AppError('Googleのクライアントが設定されていません。');
  const popup = openPopup();
  const state = randomString(24);
  const params = {
    client_id: clientId,
    redirect_uri: redirectUri(),
    response_type: 'token',
    scope: G_SCOPES,
    include_granted_scopes: 'true',
    state,
  };
  if (account) params.login_hint = account;
  const url = `${G_AUTH}?${new URLSearchParams(params)}`;
  const data = await runFlow(popup, url, state, { provider: 'google', clientId });
  return googleFinish(data);
}

async function googleFinish(data) {
  if (data.error) throw new AppError(`Googleのサインインに失敗しました：${data.error}`);
  if (!data.access_token) throw new AppError('Googleから許可を受け取れませんでした。');
  await kvSet('g_tokens', {
    accessToken: data.access_token,
    expiresAt: Date.now() + (Number(data.expires_in) || 3600) * 1000,
    scope: data.scope || '',
  });
  await kvSet('g_consented', true);
  return true;
}

export async function googleGetToken() {
  const t = await kvGet('g_tokens');
  if (!t || t.expiresAt - 60_000 <= Date.now()) throw new NeedLogin('google');
  return t.accessToken;
}

export async function googleTokenValid() {
  const t = await kvGet('g_tokens');
  return !!(t && t.expiresAt - 60_000 > Date.now());
}
export async function googleConsented() { return !!(await kvGet('g_consented')); }
export async function googleLogout() { await kvDel('g_tokens'); await kvDel('g_consented'); }

/** 画面ごと移動してサインインした場合の続き（起動時に呼ぶ） */
export async function completePendingAuth() {
  let pending;
  let result;
  try {
    pending = JSON.parse(localStorage.getItem(PENDING_KEY) || 'null');
    result = JSON.parse(localStorage.getItem(RESULT_KEY) || 'null');
  } catch { return null; }
  if (!pending || !result || !result.data || result.data.state !== pending.state) return null;
  localStorage.removeItem(PENDING_KEY);
  localStorage.removeItem(RESULT_KEY);
  if (Date.now() - pending.at > 10 * 60 * 1000) return null;
  if (pending.provider === 'ms') { await msFinish(pending.clientId, result.data, pending.verifier); return 'ms'; }
  if (pending.provider === 'google') { await googleFinish(result.data); return 'google'; }
  return null;
}
