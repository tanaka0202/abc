// サインイン（Microsoft：PKCE方式、Google：ブラウザ向けトークン方式）と、切れたサインインの自動のつなぎ直し。
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
    const onMsg = (e) => {
      if (e.origin && e.origin !== location.origin) return; // このアプリ自身（戻り先ページ）からの知らせだけを受け取る
      if (e.data && e.data.type === CHANNEL) finish(e.data.data);
    };
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

// ---------------------------------------------------------------- 自動のつなぎ直し（共通）
// ブラウザだけで動くアプリは、サービス側の決まりで、Microsoftは約24時間、Googleは約1時間でサインインが切れる。
// そこで、切れる前（または切れた直後）に、画面を出さずにつなぎ直す。方法は2つ：
//   (1) 見えない小さな枠（iframe）の中で、サービスに「画面を出さずに許可を出し直して」と頼む（prompt=none）
//   (2) (1)が使えないブラウザでは、アプリを開いた直後に、画面ごと一瞬だけサービスへ移動して戻る
// どちらも、ブラウザにサービス側のサインイン状態（Microsoft・Googleにログイン中であること）が残っている間だけ成功する。
// パスワードの入力が必要な状態では失敗し、従来どおり、利用者の操作でサインインし直す。
const SILENT_COOLDOWN = 10 * 60 * 1000;   // 失敗したら10分は試さない（何度も通信しない）
const BOUNCE_COOLDOWN = 12 * 3600 * 1000; // 画面ごとの移動に失敗したら12時間は試さない（開くたびに画面が切り替わるのを防ぐ）
const mem = { ms: { failAt: 0, busy: null }, google: { failAt: 0, busy: null } };

async function silentInfo() { return (await kvGet('auth_silent')) || {}; }
async function noteSilent(provider, patch) {
  const all = await silentInfo();
  all[provider] = { ...(all[provider] || {}), ...patch };
  await kvSet('auth_silent', all);
}

/** 見えない枠の中でサインイン画面のURLを開き、結果を待つ。時間内に結果が来なければ { error: 'timeout' } */
function silentFrame(url, state, timeoutMs) {
  const f = document.createElement('iframe');
  f.setAttribute('aria-hidden', 'true');
  f.setAttribute('tabindex', '-1');
  f.title = 'サインインの更新';
  Object.assign(f.style, { position: 'fixed', left: '-20px', top: '-20px', width: '1px', height: '1px', border: '0', opacity: '0', pointerEvents: 'none' });
  f.src = url;
  document.body.appendChild(f);
  return waitForResult(state, timeoutMs).catch(() => ({ error: 'timeout' })).finally(() => f.remove());
}

/** 画面ごと移動してのつなぎ直しを、いま試してよいか（失敗の直後や、前回の結果が不明なときは試さない） */
export async function canBounce(provider) {
  const x = (await silentInfo())[provider] || {};
  const last = x.bounceAt || 0;
  const now = Date.now();
  if (now - last < 3 * 60 * 1000) return false;
  const settled = Math.max(x.bounceOkAt || 0, x.bounceFailAt || 0) >= last;
  if (last && !settled && now - last < BOUNCE_COOLDOWN) return false; // 前回、戻ってこなかった
  if (x.bounceFailAt && now - x.bounceFailAt < BOUNCE_COOLDOWN) return false;
  return true;
}

const FRAME_RETEST = 6 * 3600 * 1000;
/**
 * 見えない枠でのつなぎ直しが、この端末で使えないと分かっているか（最近6時間のうちに失敗したか）。
 * 分かっていれば、枠は試さず、起動の直後に画面ごとの移動を選ぶ。6時間たてば、もう一度だけ枠を試す。
 */
export async function frameKnownFail(provider) {
  const x = (await silentInfo())[provider] || {};
  return x.frame === 'fail' && Date.now() - (x.frameAt || 0) < FRAME_RETEST;
}

/** サービスに届くか（オフラインのときに画面ごと移動して、エラー画面に取り残されるのを防ぐ） */
export async function reachable(provider) {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) return false;
  const url = provider === 'ms' ? `${MS_AUTHORITY}/.well-known/openid-configuration` : 'https://www.googleapis.com/calendar/v3/';
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), 3000);
  try { await fetch(url, { mode: 'no-cors', cache: 'no-store', signal: ctl.signal }); return true; } catch { return false; } finally { clearTimeout(timer); }
}

/**
 * 画面ごと一瞬だけサービスへ移動して、つなぎ直す（戻ると completePendingAuth が続きを行う）。
 * 利用者が操作を始める前（アプリを開いた直後）だけに使うこと。
 */
export async function silentRedirect(provider, { clientId, account = '' }) {
  const state = randomString(24);
  let url;
  const pending = { provider, clientId, state, at: Date.now(), silent: true };
  if (provider === 'ms') {
    const verifier = randomString(64);
    pending.verifier = verifier;
    url = await msAuthUrl(clientId, verifier, state, { prompt: 'none' });
  } else {
    url = googleAuthUrl(clientId, account, state, 'none');
  }
  localStorage.setItem(PENDING_KEY, JSON.stringify(pending));
  await noteSilent(provider, { bounceAt: Date.now() });
  location.replace(url);
  return new Promise(() => {}); // 画面が切り替わる
}

// ---------------------------------------------------------------- Microsoft
async function msAuthUrl(clientId, verifier, state, { prompt = '', hint = true } = {}) {
  const params = {
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri(),
    response_mode: 'query',
    scope: MS_SCOPES,
    code_challenge: await pkceChallenge(verifier),
    code_challenge_method: 'S256',
    state,
  };
  if (prompt) params.prompt = prompt;
  const acc = hint ? await kvGet('ms_account') : null;
  if (acc && acc.hint) params.login_hint = acc.hint;
  return `${MS_AUTHORITY}/authorize?${new URLSearchParams(params)}`;
}

/**
 * サインイン画面を開く。quick＝前回と同じアカウントで、できるだけ画面を出さずにつなぎ直す
 * （ブラウザにMicrosoftのサインイン状態が残っていれば、窓は開いてすぐ閉じる）。
 */
export async function msLogin(clientId, { quick = false } = {}) {
  if (!clientId) throw new AppError('OneDriveのアプリID（クライアントID）が設定されていません。');
  const popup = openPopup();
  const verifier = randomString(64);
  const state = randomString(24);
  const acc = await kvGet('ms_account');
  const useHint = quick && !!(acc && acc.hint);
  const url = await msAuthUrl(clientId, verifier, state, { prompt: useHint ? '' : 'select_account', hint: useHint });
  // Microsoftのサインイン画面は窓の開閉を確認できる（Googleの画面は、確認すると誤って「閉じた」と判定されることがあるため対象外）
  const data = await runFlow(popup, url, state, { provider: 'ms', verifier, clientId }, { watchClosed: true });
  await msFinish(clientId, data, verifier);
  await noteSilent('ms', { bounceFailAt: 0, bounceAt: 0, frame: '' }); // 手動でサインインし直したら、自動のつなぎ直しを最初から試す
  mem.ms.failAt = 0;
  return true;
}

/** IDトークン（サインインした人の情報）から、アカウント名を取り出す。次回のつなぎ直しで「同じアカウントで」と伝えるために使う */
function accountFromIdToken(idToken) {
  try {
    const part = String(idToken).split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
    const bytes = Uint8Array.from(atob(part + '='.repeat((4 - (part.length % 4)) % 4)), (c) => c.charCodeAt(0));
    const claims = JSON.parse(new TextDecoder().decode(bytes));
    return claims.preferred_username || claims.email || '';
  } catch { return ''; }
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
  const prev = (await kvGet('ms_account')) || {};
  await kvSet('ms_account', { hint: (j.id_token && accountFromIdToken(j.id_token)) || prev.hint || '', keep: true });
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

const MS_MAX_AGE = 23.5 * 3600_000; // ブラウザ向けアプリの更新用トークンは、サインインから約24時間で使えなくなる

let msRefreshing = null;
function msRefresh(clientId, t) {
  if (!msRefreshing) {
    msRefreshing = (async () => {
      const body = new URLSearchParams({ client_id: clientId, grant_type: 'refresh_token', refresh_token: t.refreshToken, scope: MS_SCOPES });
      const r = await fetch(`${MS_AUTHORITY}/token`, { method: 'POST', body });
      const j = await r.json().catch(() => ({}));
      if (!r.ok) {
        if (j.error === 'invalid_grant' || j.error === 'interaction_required') {
          const acc = (await kvGet('ms_account')) || {};
          await kvSet('ms_account', { ...acc, keep: true }); // サインインしていたことは覚えておく（自動でつなぎ直すため）
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

async function msSilentFrame(clientId, timeoutMs) {
  const verifier = randomString(64);
  const state = randomString(24);
  const url = await msAuthUrl(clientId, verifier, state, { prompt: 'none' });
  const data = await silentFrame(url, state, timeoutMs);
  if (data.code) {
    await msFinish(clientId, data, verifier);
    await noteSilent('ms', { frame: 'ok', frameAt: Date.now(), frameErr: '' });
    return true;
  }
  await noteSilent('ms', { frame: 'fail', frameAt: Date.now(), frameErr: String(data.error || '') });
  return false;
}

/**
 * 画面を出さずにOneDriveへつなぎ直す（見えない枠を使う）。成功したら true。
 * 以前にサインインしたことがある端末でだけ試す。失敗したら10分は試さない。
 */
export async function msSilent(clientId, { timeoutMs = 8000, force = false } = {}) {
  if (!clientId || typeof document === 'undefined') return false;
  if (!force && (Date.now() - mem.ms.failAt < SILENT_COOLDOWN || (await frameKnownFail('ms')))) return false;
  if (!(await kvGet('ms_tokens')) && !((await kvGet('ms_account')) || {}).keep) return false;
  if (!mem.ms.busy) {
    mem.ms.busy = msSilentFrame(clientId, timeoutMs).catch(() => false)
      .then((ok) => { mem.ms.failAt = ok ? 0 : Date.now(); return ok; })
      .finally(() => { mem.ms.busy = null; });
  }
  return mem.ms.busy;
}

/**
 * OneDriveを使うためのトークンを返す。期限切れなら更新し、更新できなければ画面を出さずにつなぎ直す。
 * silent:false＝つなぎ直しを待たずに、すぐ NeedLogin を返す（利用者の操作の直後に、サインイン画面を開けるようにするため）。
 */
export async function msGetToken(clientId, { silent = true } = {}) {
  let t = await kvGet('ms_tokens');
  if (t && t.expiresAt - 60_000 > Date.now()) return t.accessToken;
  const tooOld = !!(t && t.signedInAt && Date.now() - t.signedInAt > MS_MAX_AGE);
  if (t && t.refreshToken && !tooOld) {
    try { return await msRefresh(clientId, t); } catch (e) { if (!(e instanceof NeedLogin)) throw e; }
  }
  if (silent && (await msSilent(clientId))) {
    t = await kvGet('ms_tokens');
    if (t) return t.accessToken;
  }
  if (!t && !(await kvGet('ms_account'))) throw new NeedLogin('ms');
  throw new NeedLogin('ms', 'OneDriveのサインインの有効期限（約24時間）が切れました。再サインインしてください。');
}

export async function msSignedIn() { return !!(await kvGet('ms_tokens')); }
export async function msLogout() { await kvDel('ms_tokens'); await kvDel('ms_account'); }

/** サインインの残り具合：'no'＝つなぎ直し不要（または未サインイン）／'soon'＝早めに更新したい／'expired'＝切れている */
export async function msRenewal() {
  const t = await kvGet('ms_tokens');
  const acc = (await kvGet('ms_account')) || {};
  if (!t) return acc.keep ? 'expired' : 'no';
  const age = Date.now() - (t.signedInAt || 0);
  if (age > MS_MAX_AGE) return 'expired';
  return age > 6 * 3600_000 ? 'soon' : 'no';
}

// ---------------------------------------------------------------- Google
function googleAuthUrl(clientId, account, state, prompt = '') {
  const params = {
    client_id: clientId,
    redirect_uri: redirectUri(),
    response_type: 'token',
    scope: G_SCOPES,
    include_granted_scopes: 'true',
    state,
  };
  if (account) params.login_hint = account;
  if (prompt) params.prompt = prompt;
  return `${G_AUTH}?${new URLSearchParams(params)}`;
}

export async function googleLogin(clientId, account = '') {
  if (!clientId) throw new AppError('Googleのクライアントが設定されていません。');
  const popup = openPopup();
  const state = randomString(24);
  const data = await runFlow(popup, googleAuthUrl(clientId, account, state), state, { provider: 'google', clientId });
  await googleFinish(data);
  await noteSilent('google', { bounceFailAt: 0, bounceAt: 0, frame: '' });
  mem.google.failAt = 0;
  return true;
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

async function googleSilentFrame(clientId, account, timeoutMs) {
  const state = randomString(24);
  const data = await silentFrame(googleAuthUrl(clientId, account, state, 'none'), state, timeoutMs);
  if (data.access_token) {
    await googleFinish(data);
    await noteSilent('google', { frame: 'ok', frameAt: Date.now(), frameErr: '' });
    return true;
  }
  await noteSilent('google', { frame: 'fail', frameAt: Date.now(), frameErr: String(data.error || '') });
  return false;
}

/** 画面を出さずにGoogleカレンダーへつなぎ直す（一度連携した端末でだけ試す）。成功したら true */
export async function googleSilent(clientId, account = '', { timeoutMs = 8000, force = false } = {}) {
  if (!clientId || typeof document === 'undefined') return false;
  if (!(await googleConsented())) return false;
  if (!force && (Date.now() - mem.google.failAt < SILENT_COOLDOWN || (await frameKnownFail('google')))) return false;
  if (!mem.google.busy) {
    mem.google.busy = googleSilentFrame(clientId, account, timeoutMs).catch(() => false)
      .then((ok) => { mem.google.failAt = ok ? 0 : Date.now(); return ok; })
      .finally(() => { mem.google.busy = null; });
  }
  return mem.google.busy;
}

/** カレンダーを使える状態にする：有効ならそのまま、切れていれば画面を出さずにつなぎ直す。使えるなら true */
export async function googleEnsure(clientId, account = '', opts = {}) {
  if (await googleTokenValid()) return true;
  return googleSilent(clientId, account, opts);
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

/** サインインの残り具合：'no'＝更新不要（または未連携）／'soon'＝まもなく切れる／'expired'＝切れている */
export async function googleRenewal() {
  if (!(await googleConsented())) return 'no';
  const t = await kvGet('g_tokens');
  if (!t || t.expiresAt - 60_000 <= Date.now()) return 'expired';
  return t.expiresAt - Date.now() < 15 * 60_000 ? 'soon' : 'no';
}

/**
 * 画面ごと移動してサインインした場合の続き（起動時に呼ぶ）。
 * 戻り値：null、または { provider, silent, ok, error }。silent＝自動のつなぎ直し（失敗しても利用者には知らせない）
 */
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
  const provider = pending.provider;
  if (provider !== 'ms' && provider !== 'google') return null;
  const finish = () => (provider === 'ms' ? msFinish(pending.clientId, result.data, pending.verifier) : googleFinish(result.data));
  if (!pending.silent) {
    await finish();
    await noteSilent(provider, { bounceFailAt: 0, bounceAt: 0, frame: '' });
    return { provider, silent: false, ok: true };
  }
  try {
    await finish();
    await noteSilent(provider, { bounceOkAt: Date.now(), bounceFailAt: 0 });
    mem[provider].failAt = 0;
    return { provider, silent: true, ok: true };
  } catch (e) {
    await noteSilent(provider, { bounceFailAt: Date.now(), bounceErr: String(result.data.error || e.message || '') });
    return { provider, silent: true, ok: false, error: String(result.data.error || '') };
  }
}
