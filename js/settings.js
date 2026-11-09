// 設定（端末内に保存。秘密でない項目はOneDrive経由で2台に共有）
import { kvGet, kvSet } from './db.js';
import { deepMerge, uuid, defaultDeviceName } from './util.js';

export const DEFAULT_SETTINGS = {
  deviceId: '',
  deviceName: '',
  gemini: {
    apiKey: '',
    paidConfirmed: false,          // 課金を有効にしたプロジェクトのキーであることを利用者が確認したか
    transcribeModel: 'gemini-3.5-transcribe',
    textModel: 'gemini-3.8-flash', // 用語補正・単語抽出
    minutesModel: 'gemini-3.8-flash',
    language: 'ja-JP',
    chunkMinutes: 20,              // 話者分離ありは1回30分までなので余裕を持たせる
  },
  onedrive: {
    clientId: '',
    rootPath: '会議録音アーカイブ',
  },
  google: {
    clientId: '',
    account: '',                   // 例: 自分のGmailアドレス（ログインのヒント）
    calendarId: 'primary',
    autoCalendar: true,
  },
  recording: {
    highpass: true,
    noiseSuppression: true,
    autoGainControl: true,
    echoCancellation: false,
    holdToStop: true,
    autoStep1: true,               // 録音終了後に自動で①を実行
  },
  transcript: {
    fillerLevel: 'standard',       // 'none' | 'standard'
    diarization: true,
  },
  memo: {
    phrases: ['決定', '要確認', '次回までに対応', '資料共有待ち', '要再検討', '異議なし'],
  },
  sync: {
    enabled: true,
    pollSec: 5,
  },
  privacy: {
    localRetentionDays: 7,
    calendarSummary: true,         // カレンダーに要約を書く（会議ごとの「機密」で個別に無効化）
  },
  budget: {
    monthlyUsd: 10,
  },
  lastProjectId: 'common',
  lastTemplateId: 'standard',
};

// 2台で共有してよい項目（APIキー・トークン・端末情報は含めない）
const SHARED_KEYS = [
  ['gemini', 'transcribeModel'], ['gemini', 'textModel'], ['gemini', 'minutesModel'], ['gemini', 'language'], ['gemini', 'chunkMinutes'],
  ['google', 'clientId'], ['google', 'account'], ['google', 'calendarId'], ['google', 'autoCalendar'],
  ['transcript', 'fillerLevel'], ['transcript', 'diarization'],
  ['memo', 'phrases'], ['privacy', 'calendarSummary'], ['budget', 'monthlyUsd'],
];

let current = null;
const listeners = new Set();

export async function loadSettings() {
  const saved = await kvGet('settings', {});
  current = deepMerge(DEFAULT_SETTINGS, saved || {});
  let changed = false;
  if (!current.deviceId) { current.deviceId = uuid().slice(0, 8); changed = true; }
  if (!current.deviceName) { current.deviceName = defaultDeviceName(); changed = true; }
  if (changed) await kvSet('settings', current);
  return current;
}

export function getSettings() {
  if (!current) throw new Error('設定が読み込まれていません');
  return current;
}

export async function saveSettings(patch) {
  current = deepMerge(current || DEFAULT_SETTINGS, patch);
  await kvSet('settings', current);
  for (const fn of listeners) { try { fn(current); } catch (e) { console.error(e); } }
  return current;
}

export function onSettingsChange(fn) { listeners.add(fn); return () => listeners.delete(fn); }

export function sharedPart(s = current) {
  const out = {};
  for (const [a, b] of SHARED_KEYS) {
    if (s[a] && b in s[a]) { out[a] = out[a] || {}; out[a][b] = s[a][b]; }
  }
  return out;
}

/** OneDrive上の共有設定を取り込む（端末固有・秘密の項目は無視） */
export async function applyShared(shared) {
  if (!shared || typeof shared !== 'object') return current;
  const patch = {};
  for (const [a, b] of SHARED_KEYS) {
    if (shared[a] && b in shared[a]) { patch[a] = patch[a] || {}; patch[a][b] = shared[a][b]; }
  }
  return saveSettings(patch);
}

/** 設定状況のチェック（初回設定の案内に使う） */
export function setupStatus(s = current) {
  return {
    gemini: !!(s.gemini.apiKey && s.gemini.paidConfirmed),
    geminiKeyOnly: !!(s.gemini.apiKey && !s.gemini.paidConfirmed),
    onedrive: !!s.onedrive.clientId,
    google: !!s.google.clientId,
  };
}
