// 汎用ユーティリティ（ブラウザ・Node両方で使える純粋関数を中心に）

/** 秒 → "MM:SS"（1時間以上は "H:MM:SS"） */
export function fmtTime(sec) {
  if (!Number.isFinite(sec) || sec < 0) sec = 0;
  const s = Math.floor(sec);
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const r = s % 60;
  const mm = String(m).padStart(2, '0');
  const ss = String(r).padStart(2, '0');
  return h > 0 ? `${h}:${mm}:${ss}` : `${mm}:${ss}`;
}

/** "MM:SS" / "H:MM:SS" / "[MM:SS]" → 秒。解釈できなければ null */
export function parseTimeTag(tag) {
  if (typeof tag !== 'string') return null;
  const m = tag.trim().replace(/^\[|\]$/g, '').match(/^(?:(\d+):)?(\d{1,3}):(\d{2})$/);
  if (!m) return null;
  const h = m[1] ? Number(m[1]) : 0;
  return h * 3600 + Number(m[2]) * 60 + Number(m[3]);
}

/** "1.500s" / "1.5" / {seconds, nanos} → 秒 */
export function parseDuration(v) {
  if (v == null) return null;
  if (typeof v === 'number') return v;
  if (typeof v === 'object') {
    const s = Number(v.seconds || 0);
    const n = Number(v.nanos || 0);
    return s + n / 1e9;
  }
  const m = String(v).trim().match(/^(-?\d+(?:\.\d+)?)s?$/);
  return m ? Number(m[1]) : null;
}

/** ローカル時刻で YYYYMMDD_HHMM */
export function folderNameFor(date) {
  const d = new Date(date);
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

/** 表示用 YYYY/MM/DD HH:MM */
export function fmtDateTime(date) {
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) return '';
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function uuid() {
  if (globalThis.crypto && crypto.randomUUID) return crypto.randomUUID();
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    return (c === 'x' ? r : (r & 0x3) | 0x8).toString(16);
  });
}

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export function debounce(fn, ms) {
  let t = null;
  const wrapped = (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
  wrapped.flush = (...args) => { clearTimeout(t); return fn(...args); };
  return wrapped;
}

/** Blob → base64（データURLの接頭辞なし） */
export function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const s = String(fr.result);
      resolve(s.slice(s.indexOf(',') + 1));
    };
    fr.onerror = () => reject(fr.error || new Error('読み込みに失敗しました'));
    fr.readAsDataURL(blob);
  });
}

/** ファイル名に使えない文字を置換 */
export function safeFileName(name) {
  return String(name || '').replace(/[\\/:*?"<>|#%\u0000-\u001f]/g, '_').trim().slice(0, 120) || 'file';
}

/** 簡易な深いマージ（配列は上書き） */
export function deepMerge(base, over) {
  if (Array.isArray(base) || Array.isArray(over)) return over === undefined ? base : over;
  if (typeof base !== 'object' || base === null) return over === undefined ? base : over;
  const out = { ...base };
  if (typeof over === 'object' && over !== null) {
    for (const k of Object.keys(over)) {
      out[k] = k in base ? deepMerge(base[k], over[k]) : over[k];
    }
  }
  return out;
}

/** JSON を安全に取り出す（前後に余計な文字があっても {...} / [...] を探す） */
export function parseJsonLoose(text) {
  if (typeof text !== 'string') return null;
  const t = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  try { return JSON.parse(t); } catch { /* 続行 */ }
  const starts = [t.indexOf('{'), t.indexOf('[')].filter((i) => i >= 0);
  if (!starts.length) return null;
  const start = Math.min(...starts);
  const end = Math.max(t.lastIndexOf('}'), t.lastIndexOf(']'));
  if (end <= start) return null;
  try { return JSON.parse(t.slice(start, end + 1)); } catch { return null; }
}

export function isAndroid() { return typeof navigator !== 'undefined' && /Android/i.test(navigator.userAgent); }
export function isIOS() {
  if (typeof navigator === 'undefined') return false;
  return /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
}
export function isMobile() { return isAndroid() || isIOS(); }

export function defaultDeviceName() {
  if (isAndroid()) return 'Android';
  if (isIOS()) return 'iPad/iPhone';
  if (typeof navigator !== 'undefined' && /Windows/i.test(navigator.userAgent)) return 'Windows';
  if (typeof navigator !== 'undefined' && /Mac/i.test(navigator.userAgent)) return 'Mac';
  return 'PC';
}

/** 例外を利用者向けの日本語に */
export function errorMessage(e) {
  if (!e) return '不明なエラー';
  if (typeof e === 'string') return e;
  return e.userMessage || e.message || String(e);
}

export class AppError extends Error {
  constructor(userMessage, { code, cause } = {}) {
    super(userMessage);
    this.userMessage = userMessage;
    this.code = code;
    this.cause = cause;
  }
}

/** DOM生成ヘルパー（文字列はすべて textContent として入れる＝HTMLとして解釈しない） */
export function el(tag, attrs = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'dataset') Object.assign(node.dataset, v);
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'value') node.value = v;
    else if (k === 'checked') node.checked = !!v;
    else if (k === 'disabled') node.disabled = !!v;
    else node.setAttribute(k, v === true ? '' : String(v));
  }
  appendChildren(node, children);
  return node;
}

function appendChildren(node, children) {
  for (const c of children.flat(Infinity)) {
    if (c === null || c === undefined || c === false) continue;
    node.appendChild(c instanceof Node ? c : document.createTextNode(String(c)));
  }
}

export function clear(node) { while (node && node.firstChild) node.removeChild(node.firstChild); return node; }

export function $(sel, root = document) { return root.querySelector(sel); }
export function $$(sel, root = document) { return Array.from(root.querySelectorAll(sel)); }
