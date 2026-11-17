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
    chunkMinutes: 10,              // 区切りの長さ。話者分離は1回30分まで。スマホのメモリを考えて10分
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
    autoSave: false,               // false＝録音終了後、ボタンを押してから①（保存・書き起こし）を始める
    darkOnStart: 'auto',           // 録音開始で画面を黒くする：'auto'（スマホ・タブレットのみ）|'on'|'off'
    audioFormat: 'm4a',            // OneDriveに保存する録音データの形式：'m4a'（AACで圧縮。小さい）|'wav'（無圧縮。大きい）
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
    // 書き起こしを行う端末：'auto'＝スマホ・タブレットは保存だけ行い、PCが書き起こす／'yes'＝この端末で書き起こす／'no'＝この端末では書き起こさない
    transcribe: 'auto',
  },
  connection: {
    autoRenew: true,               // OneDrive・Googleカレンダーのサインインを、切れる前に自動でつなぎ直す
  },
  minutes: {
    auto: false,                   // true＝書き起こしが終わったら、続けて議事録（②）まで自動で作る
  },
  privacy: {
    localRetentionDays: 7,
    calendarSummary: true,         // カレンダーに要約を書く（会議ごとの「機密」で個別に無効化）
  },
  budget: {
    monthlyUsd: 10,
  },
  lastScene: 'room',             // 録音の場面（room：対面、speaker：PCのスピーカーの音、tab：PCのタブ音声）
  lastProjectId: 'common',
  lastTemplateId: 'standard',
};

// 2台で共有してよい項目（APIキー・トークン・端末情報は含めない）
const SHARED_KEYS = [
  ['gemini', 'transcribeModel'], ['gemini', 'textModel'], ['gemini', 'minutesModel'], ['gemini', 'language'], ['gemini', 'chunkMinutes'],
  ['google', 'clientId'], ['google', 'account'], ['google', 'calendarId'], ['google', 'autoCalendar'],
  ['transcript', 'fillerLevel'], ['transcript', 'diarization'],
  ['memo', 'phrases'], ['privacy', 'calendarSummary'], ['budget', 'monthlyUsd'], ['minutes', 'auto'],
];

let current = null;
const listeners = new Set();

export async function loadSettings() {
  const saved = await kvGet('settings', {});
  current = deepMerge(DEFAULT_SETTINGS, saved || {});
  let changed = false;
  // 旧版の設定の引き継ぎ：「黒くしない」を選んでいた場合だけ引き継ぐ。自動で①を始める設定は引き継がない
  const oldRec = (saved && saved.recording) || {};
  if (oldRec.blackout === false && oldRec.darkOnStart === undefined) { current.recording.darkOnStart = 'off'; changed = true; }
  if ('blackout' in current.recording || 'autoStep1' in current.recording) { delete current.recording.blackout; delete current.recording.autoStep1; changed = true; }
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
