// 画面全体の制御
import { el, clear, $, fmtTime, fmtDateTime, uuid, errorMessage, isMobile, debounce, sleep } from './util.js';
import { loadSettings, getSettings, saveSettings, setupStatus, sharedPart, applyShared } from './settings.js';
import { putSession, getSession, requestPersistence, wipeAll, getChunks, getFiles, kvGet, kvSet } from './db.js';
import { Recorder } from './recorder.js';
import { msLogin, msGetToken, msSignedIn, msLogout, googleLogin, googleTokenValid, googleLogout, completePendingAuth, NeedLogin, redirectUri } from './auth.js';
import * as od from './onedrive.js';
import * as vocab from './vocab.js';
import * as memos from './memos.js';
import * as pipe from './pipeline.js';
import { LivePublisher, LiveWatcher, requestRemoteStop, readAllMemos, rootPath, metaPath } from './sync.js';
import { testApiKey } from './gemini.js';
import { eventsAround } from './gcal.js';
import * as tp from './textproc.js';
import { getUsage, getKaizen } from './usage.js';
import { pcmBlobFromChunks, wavFromPcmBlob } from './wav.js';
import { qrCanvas } from './qr.js';
import { SETUP_PREFIX, setupPayload, encodeSetup, parseSetup } from './setuplink.js';
import { toast, holdButton, confirmDialog, formDialog, renderMarkdown, chip, progressBar, section } from './ui.js';

// ---------------------------------------------------------------------------
// 状態
// ---------------------------------------------------------------------------
const S = {
  view: 'meeting',
  rec: null,            // 録音中の Recorder
  session: null,        // この端末の現在の会議
  publisher: null,      // 録音状況の発信
  remote: null,         // 他の端末で録音中・処理中の会議
  remoteMeta: null,     // その会議の session.json
  draft: null,          // 録音前に書いたメモの受け皿
  memoT0: null,         // メモを書き始めた時刻
  myMemos: [],
  otherMemos: [],
  proc: null,           // 処理中の状態 { sessionId, kind, progress }
  detail: null,         // 詳細表示中の会議
  detailTab: 'transcript',
  audioUrl: null,
  editTranscript: false,
  editMinutes: false,
  list: [],
  listQuery: '',
  searchHits: null,
  vocabProject: 'common',
  recovered: [],
  warnings: [],
  status: { ms: false, google: false },
  level: -100,
};
const watcher = new LiveWatcher({ onChange: onRemoteChange, onError: () => {} });

// ---------------------------------------------------------------------------
// 起動
// ---------------------------------------------------------------------------
async function boot() {
  await loadSettings();
  await vocab.loadVocab();
  requestPersistence();
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js', { updateViaCache: 'none' }).catch(() => {});

  // ほかの端末のQRコードから開かれた場合（設定は画面の中だけで扱い、URLからはすぐ消す）
  let setupRaw = null;
  if (location.hash.startsWith(SETUP_PREFIX)) {
    setupRaw = location.hash.slice(SETUP_PREFIX.length);
    history.replaceState(null, '', location.pathname + location.search);
  }
  if (location.hash === '#auth-return') {
    history.replaceState(null, '', location.pathname + location.search);
    try { const p = await completePendingAuth(); if (p) toast(p === 'ms' ? 'OneDriveにサインインしました' : 'Googleカレンダーを連携しました', 'ok'); } catch (e) { toast(errorMessage(e), 'error'); }
  }
  S.status.ms = await msSignedIn();
  S.status.google = await googleTokenValid();
  S.vocabProject = getSettings().lastProjectId || 'common';
  vocab.setVocabUploader(uploadVocab);

  S.recovered = await pipe.findInterrupted(null);
  const rt = await kvGet('remoteTarget');
  if (rt && Date.now() - (rt.savedAt || 0) < 24 * 3600 * 1000) S.remote = rt;

  bindNav();
  render();
  if (setupRaw) offerSetupImport(setupRaw);
  // アプリを開いたまま、同じ画面でQRコードのリンクが開かれた場合（ページは読み込み直されない）
  window.addEventListener('hashchange', () => {
    if (!location.hash.startsWith(SETUP_PREFIX)) return;
    const raw = location.hash.slice(SETUP_PREFIX.length);
    history.replaceState(null, '', location.pathname + location.search);
    offerSetupImport(raw);
  });
  if (new URLSearchParams(location.search).get('action') === 'record') setTimeout(() => { const b = $('#btn-start'); if (b) b.focus(); }, 300);

  window.addEventListener('beforeunload', (e) => { if (S.rec || S.proc) { e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') refreshRemoteMemos(); });

  if (S.status.ms) afterMsReady();
  pipe.cleanupLocal().catch(() => {});
  setInterval(refreshRemoteMemos, 15000);
  setInterval(pollRemoteMeta, 10000);
  setInterval(tickRemote, 1000);
}

async function afterMsReady() {
  if (getSettings().sync.enabled) watcher.start();
  // 単語帳・設定を2台で共有
  try {
    const remoteVocab = await od.readJson(od.joinPath(rootPath(), '_app', 'vocabulary.json'));
    if (remoteVocab) { const changed = await vocab.mergeRemote(remoteVocab); if (changed) await uploadVocab(vocab.getVocab()); }
    else await uploadVocab(vocab.getVocab());
    const shared = await od.readJson(od.joinPath(rootPath(), '_app', 'settings.json'));
    if (shared) await applyShared(shared);
    else await od.uploadJson(od.joinPath(rootPath(), '_app', 'settings.json'), sharedPart());
  } catch (e) {
    if (e instanceof NeedLogin) { S.status.ms = false; renderHeader(); }
  }
  if (S.view === 'list') loadList();
}

async function uploadVocab(v) {
  if (!getSettings().onedrive.clientId || !S.status.ms) return;
  await od.ensureFolder(od.joinPath(rootPath(), '_app'));
  await od.uploadJson(od.joinPath(rootPath(), '_app', 'vocabulary.json'), v);
}

async function ensureMs(interactive) {
  const st = getSettings();
  if (!st.onedrive.clientId) return false;
  try {
    await msGetToken(st.onedrive.clientId);
    if (!S.status.ms) { S.status.ms = true; renderHeader(); }
    return true;
  } catch (e) {
    if (e instanceof NeedLogin && interactive) {
      await msLogin(st.onedrive.clientId);
      S.status.ms = true;
      renderHeader();
      afterMsReady();
      return true;
    }
    S.status.ms = false;
    renderHeader();
    if (!(e instanceof NeedLogin)) throw e;
    return false;
  }
}

// ---------------------------------------------------------------------------
// ナビゲーション
// ---------------------------------------------------------------------------
const VIEWS = [
  { id: 'meeting', label: '会議', icon: '●' },
  { id: 'list', label: '一覧', icon: '☰' },
  { id: 'vocab', label: '単語帳', icon: '📖' },
  { id: 'settings', label: '設定', icon: '⚙' },
];

function bindNav() {
  const nav = $('#nav');
  clear(nav);
  for (const v of VIEWS) {
    nav.appendChild(el('button', { type: 'button', class: 'nav-btn', dataset: { view: v.id }, onclick: () => go(v.id) },
      el('span', { class: 'nav-icon', 'aria-hidden': 'true' }, v.icon), el('span', {}, v.label)));
  }
}

function go(view) {
  if (S.view === 'detail' && view !== 'detail') leaveDetail();
  S.view = view;
  render();
  window.scrollTo(0, 0);
  if (view === 'list') loadList();
}

function render() {
  renderHeader();
  for (const b of document.querySelectorAll('.nav-btn')) b.classList.toggle('active', b.dataset.view === S.view || (S.view === 'detail' && b.dataset.view === 'list'));
  const main = $('#main');
  clear(main);
  if (S.view === 'meeting') main.appendChild(viewMeeting());
  else if (S.view === 'list') main.appendChild(viewList());
  else if (S.view === 'vocab') main.appendChild(viewVocab());
  else if (S.view === 'settings') main.appendChild(viewSettings());
  else if (S.view === 'detail') main.appendChild(viewDetail());
}

function renderHeader() {
  const box = $('#status');
  if (!box) return;
  clear(box);
  const st = getSettings();
  const ss = setupStatus(st);
  box.appendChild(chip(!ss.onedrive ? 'OneDrive 未設定' : S.status.ms ? 'OneDrive ✓' : 'OneDrive 要サインイン', !ss.onedrive ? 'muted' : S.status.ms ? 'ok' : 'warn'));
  box.appendChild(chip(ss.gemini ? 'Gemini ✓' : ss.geminiKeyOnly ? 'Gemini 課金確認待ち' : 'Gemini 未設定', ss.gemini ? 'ok' : 'warn'));
  if (ss.google) box.appendChild(chip(S.status.google ? 'カレンダー ✓' : 'カレンダー 要更新', S.status.google ? 'ok' : 'muted'));
  if (S.rec) box.appendChild(chip(`録音中 ${fmtTime(S.rec.elapsed)}`, 'rec'));
  else if (S.remote && S.remote.status === 'recording') box.appendChild(chip(`${S.remote.deviceName}で録音中`, 'rec'));
}

// ===========================================================================
// 会議（録音＋メモ）
// ===========================================================================
function memoTarget() {
  if (S.rec && S.session) return S.session;
  if (S.remote) return { id: S.remote.sessionId, startedAt: S.remote.startedAt, folderPath: S.remote.folderPath, title: S.remote.title, isRemote: true };
  if (S.session && S.session.status !== 'done') return S.session;
  if (!S.draft) S.draft = { id: uuid(), startedAt: null, folderPath: null, title: '', isDraft: true };
  return S.draft;
}

function viewMeeting() {
  const wrap = el('div', { class: 'grid-2' });
  const left = el('div', { class: 'col' });
  const right = el('div', { class: 'col' });
  left.appendChild(el('div', { id: 'recover-box' }));
  left.appendChild(el('div', { id: 'remote-box' }));
  left.appendChild(el('div', { id: 'rec-box' }));
  left.appendChild(el('div', { id: 'proc-box' }));
  right.appendChild(el('div', { id: 'memo-box' }));
  wrap.append(left, right);
  queueMicrotask(() => { renderRecover(); renderRemote(); renderRec(); renderProc(); renderMemoBox(); });
  return wrap;
}

function renderRecover() {
  const box = $('#recover-box');
  if (!box) return;
  clear(box);
  for (const s of S.recovered) {
    box.appendChild(el('div', { class: 'card notice' },
      el('p', {}, `前回の録音「${s.title}」（${fmtTime(s.durationSec)}）が途中で終わっています。録音された部分は端末内に残っています。`),
      el('div', { class: 'row' },
        el('button', { class: 'btn primary', type: 'button', onclick: () => { S.recovered = S.recovered.filter((x) => x.id !== s.id); S.session = s; renderRecover(); runStep1UI(s); } }, '保存して書き起こす'),
        el('button', { class: 'btn ghost', type: 'button', onclick: () => { S.recovered = S.recovered.filter((x) => x.id !== s.id); renderRecover(); } }, '後で（一覧から操作）'))));
  }
}

// ---- 他の端末の録音（B-7） ----
function onRemoteChange(live) {
  if (S.rec) return; // 自分が録音中は無視
  if (live) {
    if (!S.remote || S.remote.sessionId !== live.sessionId) { S.otherMemos = []; S.remoteMeta = null; }
    S.remote = { ...live, savedAt: Date.now() };
    kvSet('remoteTarget', S.remote);
    if (live.status === 'recording') toast(`${live.deviceName}で録音中の会議「${live.title}」につながりました。ここで書いたメモは録音の時刻と結びつきます。`, 'ok', 6000);
  }
  renderHeader();
  if (S.view === 'meeting') { renderRemote(); renderRec(); renderMemoBox(); }
  refreshRemoteMemos();
}

function tickRemote() {
  if (S.remote && S.remote.status === 'recording' && S.view === 'meeting') {
    const t = $('#remote-time');
    if (t) t.textContent = fmtTime((Date.now() - S.remote.startedAt) / 1000);
  }
}

async function pollRemoteMeta() {
  if (!S.remote || !S.remote.folderPath || !S.status.ms || document.visibilityState !== 'visible') return;
  if (S.remoteMeta && S.remoteMeta.status === 'done') return;
  try {
    const meta = await od.readJson(metaPath(S.remote.folderPath, 'session.json'));
    if (meta && JSON.stringify(meta) !== JSON.stringify(S.remoteMeta)) {
      const was = S.remoteMeta && S.remoteMeta.status;
      S.remoteMeta = meta;
      if (meta.status === 'transcribed' && was !== 'transcribed') toast('書き起こしができました。「② 議事録作成」を押せます。', 'ok', 6000);
      if (S.view === 'meeting') renderRemote();
    }
  } catch { /* 次回 */ }
}

function closeRemote() {
  S.remote = null;
  S.remoteMeta = null;
  S.otherMemos = [];
  kvSet('remoteTarget', null);
  render();
}

function renderRemote() {
  const box = $('#remote-box');
  if (!box) return;
  clear(box);
  const r = S.remote;
  if (!r || S.rec) return;
  const meta = S.remoteMeta;
  const card = el('section', { class: 'card remote' });
  if (r.status === 'recording') {
    card.append(
      el('div', { class: 'remote-head' }, el('span', { class: 'dot rec' }), el('strong', {}, `${r.deviceName}で録音中`), el('span', { id: 'remote-time', class: 'mono big' }, fmtTime((Date.now() - r.startedAt) / 1000))),
      el('p', {}, `会議：${r.title}`),
      el('p', { class: 'muted small' }, '右（スマホでは下）のメモ欄に書くと、録音の時刻と自動で結びつきます。'),
      holdButton(`${r.deviceName}の録音を終了して①を実行`, () => remoteStop(r), { className: 'danger wide', enabled: getSettings().recording.holdToStop, hint: '長押しで終了' }));
  } else {
    const status = meta ? statusLabel(meta) : (r.status === 'lost' ? '録音端末からの連絡が途絶えています' : '録音終了・処理中');
    card.append(
      el('div', { class: 'remote-head' }, el('span', { class: 'dot' }), el('strong', {}, `会議「${r.title}」`)),
      el('p', {}, status),
      meta && meta.progress ? el('p', { class: 'muted small' }, meta.progress.text) : null,
      el('div', { class: 'row' },
        el('button', { class: 'btn primary', type: 'button', disabled: !(meta && meta.steps && meta.steps.transcript) || !!S.proc, onclick: () => runStep2UI({ id: r.sessionId, folderPath: r.folderPath }) }, '② メモを組み込んで議事録作成'),
        el('button', { class: 'btn', type: 'button', onclick: () => saveMemosUI({ id: r.sessionId, folderPath: r.folderPath, title: r.title }) }, '📌 メモだけ保存'),
        el('button', { class: 'btn ghost', type: 'button', onclick: () => openDetailByFolder(r.folderPath) }, '詳細'),
        el('button', { class: 'btn ghost', type: 'button', onclick: closeRemote }, 'この会議を閉じる')));
  }
  box.appendChild(card);
}

async function remoteStop(r) {
  try {
    await ensureMs(true);
    await requestRemoteStop(r);
    toast(`${r.deviceName}に停止を指示しました（数秒で止まります）`, 'ok');
  } catch (e) { toast(errorMessage(e), 'error'); }
}

function statusLabel(s) {
  const map = { recording: '録音中', recorded: '録音済み（①未実行）', processing: '①処理中', transcribed: '書き起こし済み（②未実行）', done: '議事録作成済み', error: 'エラー' };
  return map[s.status] || s.status || '';
}

// ---- 録音操作 ----
function renderRec() {
  const box = $('#rec-box');
  if (!box) return;
  clear(box);
  const st = getSettings();
  if (S.rec) {
    const s = S.session;
    box.appendChild(el('section', { class: 'card rec-card recording' },
      el('div', { class: 'rec-head' }, el('span', { class: 'dot rec' }), el('strong', {}, '録音中'), el('span', { id: 'rec-time', class: 'mono big' }, fmtTime(S.rec.elapsed))),
      el('div', { class: 'meter', 'aria-label': '音量' }, el('div', { id: 'meter-fill', class: 'meter-fill' })),
      el('p', { class: 'muted small' }, `会議：${s.title}${s.folderName ? `（保存先：${s.folderName}）` : ''}`),
      el('div', { id: 'rec-warn' }, S.warnings.slice(-2).map((w) => el('p', { class: 'warn small' }, w))),
      el('div', { class: 'row' },
        el('button', { class: 'btn', type: 'button', onclick: () => quickMemo('star') }, '★ ここをピン留め')),
      holdButton('① 録音を終了して保存・書き起こし', () => handleStop(false), { className: 'danger wide big', enabled: st.recording.holdToStop, hint: '長押しで終了' }),
      el('p', { class: 'muted small' }, '録音中は画面をつけたまま、このアプリを表示しておいてください（画面のスリープは自動で防いでいます）。')));
    updateMeter();
    return;
  }
  // 他の端末で録音中のときは、この端末の録音欄を小さくまとめる（押し間違い防止）
  if (S.remote && S.remote.status === 'recording' && !S.showLocalRec) {
    box.appendChild(el('section', { class: 'card' },
      el('p', { class: 'muted small' }, `${S.remote.deviceName}で録音中のため、この端末ではメモを取れます。`),
      el('button', { class: 'btn ghost small', type: 'button', onclick: () => { S.showLocalRec = true; renderRec(); } }, 'この端末でも別に録音する')));
    return;
  }
  // 録音前の設定
  const projects = vocab.projects();
  const title = el('input', { class: 'input', id: 'in-title', placeholder: '会議名（空欄なら日時）', value: S.setupTitle || '', oninput: (e) => { S.setupTitle = e.target.value; } });
  const proj = el('select', { class: 'input', id: 'in-project' }, projects.map((p) => el('option', { value: p.id, selected: p.id === st.lastProjectId }, p.name)));
  const tpl = el('select', { class: 'input', id: 'in-template' }, tp.TEMPLATES.map((t) => el('option', { value: t.id, selected: t.id === st.lastTemplateId }, t.name)));
  const people = el('input', { class: 'input', id: 'in-people', placeholder: '参加者（読点「、」区切り。話者名の推定に使います）', value: S.setupPeople || '', oninput: (e) => { S.setupPeople = e.target.value; } });
  const conf = el('label', { class: 'check' }, el('input', { type: 'checkbox', id: 'in-conf', checked: !!S.setupConf, onchange: (e) => { S.setupConf = e.target.checked; } }), '機密の会議（カレンダーに要約を書かない）');
  const diar = el('label', { class: 'check' }, el('input', { type: 'checkbox', id: 'in-diar', checked: st.transcript.diarization !== false }), '話者を分けて書き起こす');
  const scenes = [
    ['room', '対面の会議（マイクで直接録る）'],
    ['speaker', 'PCのスピーカーから流れるWeb会議を録る'],
    ...(isMobile() ? [] : [['tab', 'PCのタブ音声＋マイク（Web会議を直接録音）']]),
  ];
  const lastScene = scenes.some(([v]) => v === st.lastScene) ? st.lastScene : 'room';
  const source = el('select', { class: 'input', id: 'in-scene' }, scenes.map(([v, label]) => el('option', { value: v, selected: v === lastScene }, label)));
  const fromCal = st.google.clientId ? el('button', { class: 'btn ghost small', type: 'button', onclick: pickFromCalendar }, '📅 予定から取り込む') : null;
  const startBtn = el('button', { class: 'btn rec-btn', id: 'btn-start', type: 'button', onclick: handleStart }, el('span', { class: 'dot' }), '録音開始');
  box.appendChild(el('section', { class: 'card rec-card' },
    el('h2', {}, '録音'),
    el('div', { class: 'field' }, el('span', {}, '会議名'), el('div', { class: 'row nowrap' }, title, fromCal)),
    el('div', { class: 'row' }, el('label', { class: 'field grow' }, el('span', {}, '案件（単語帳）'), proj), el('label', { class: 'field grow' }, el('span', {}, '議事録の種類'), tpl)),
    el('label', { class: 'field' }, el('span', {}, '参加者'), people),
    el('label', { class: 'field' }, el('span', {}, '録音の場面'), source,
      el('small', { class: 'muted' }, '「PCのスピーカー」を選ぶと、離れたスピーカーの声を雑音と誤って消さないよう、雑音抑制を自動で切ります。')),
    el('div', { class: 'row' }, conf, diar),
    startBtn,
    setupHint()));
}

function setupHint() {
  const ss = setupStatus();
  const msgs = [];
  if (!ss.onedrive) msgs.push('OneDriveが未設定のため、録音は端末内だけに保存されます。');
  else if (!S.status.ms) msgs.push('録音開始時にOneDriveへのサインインを求めます。');
  if (!ss.gemini) msgs.push('Geminiが未設定（または課金の確認待ち）のため、書き起こしは設定後に行います。');
  if (!msgs.length) return null;
  return el('div', { class: 'hint' }, msgs.map((m) => el('p', { class: 'small' }, m)), el('button', { class: 'btn ghost small', type: 'button', onclick: () => go('settings') }, '設定を開く'));
}

async function pickFromCalendar() {
  try {
    if (!(await googleTokenValid())) await googleLogin(getSettings().google.clientId, getSettings().google.account);
    S.status.google = true;
    const evs = await eventsAround(new Date());
    if (!evs.length) { toast('前後2時間に予定が見つかりませんでした'); return; }
    const dlg = el('dialog', { class: 'dialog' }, el('h3', {}, '予定から取り込む'),
      el('div', { class: 'list' }, evs.map((e) => el('button', {
        class: 'list-item', type: 'button',
        onclick: () => { S.setupTitle = e.title; S.setupPeople = e.attendees.join('、'); dlg.close(); renderRec(); },
      }, el('strong', {}, e.title), el('span', { class: 'muted small' }, `${fmtDateTime(e.start)}  ${e.attendees.join('、')}`)))),
      el('div', { class: 'row end' }, el('button', { class: 'btn ghost', type: 'button', onclick: () => dlg.close() }, '閉じる')));
    dlg.addEventListener('close', () => dlg.remove());
    document.body.appendChild(dlg);
    dlg.showModal();
  } catch (e) { toast(errorMessage(e), 'error'); }
}

async function handleStart() {
  if (S.rec) return;
  const st = getSettings();
  const btn = $('#btn-start');
  if (btn) btn.disabled = true;
  try {
    let msOk = false;
    if (st.onedrive.clientId) {
      try { msOk = await ensureMs(true); } catch (e) { toast(`OneDriveに接続できませんでした。端末内に録音します（${errorMessage(e)}）`, 'warn', 6000); }
    }
    const projectId = ($('#in-project') || {}).value || 'common';
    const templateId = ($('#in-template') || {}).value || 'standard';
    const participants = String(S.setupPeople || '').split(/[、,，\s]+/).map((x) => x.trim()).filter(Boolean);
    const scene = ($('#in-scene') || {}).value || 'room';
    // 場面ごとのマイクの音声処理：PCのスピーカーの音は雑音抑制・エコー除去で消されやすいので切る
    const audioOverrides = scene === 'speaker' ? { noiseSuppression: false, echoCancellation: false, autoGainControl: true } : {};
    const draft = S.draft;
    const session = {
      id: draft ? draft.id : uuid(),
      title: (S.setupTitle || '').trim() || `会議 ${fmtDateTime(Date.now())}`,
      projectId, templateId, participants,
      confidential: !!S.setupConf,
      diarization: ($('#in-diar') || { checked: true }).checked,
      source: scene === 'tab' ? 'tab' : 'mic',
      scene,
      status: 'recording', createdAt: Date.now(), steps: {},
      deviceId: st.deviceId, deviceName: st.deviceName, materials: [],
    };
    const proj = vocab.projectById(projectId);
    session.projectName = proj ? proj.name : '';
    await putSession(session);
    S.draft = null;
    S.warnings = [];
    const rec = new Recorder({
      session, settings: st,
      onLevel: (db) => { S.level = db; updateMeter(); },
      onTick: (sec) => { const t = $('#rec-time'); if (t) t.textContent = fmtTime(sec); renderHeaderThrottled(); },
      onWarning: (code, msg) => { S.warnings.push(msg); const w = $('#rec-warn'); if (w) { clear(w); S.warnings.slice(-2).forEach((m) => w.appendChild(el('p', { class: 'warn small' }, m))); } if (code !== 'hidden') toast(msg, 'warn', 6000); },
      onStarted: async (t0) => {
        session.startedAt = t0;
        await putSession(session);
        await memos.fixOffsets(session);
        if (msOk) startOnline(session);
      },
    });
    await rec.start(session.source, audioOverrides);
    S.rec = rec;
    S.session = session;
    S.remote = null; S.remoteMeta = null; S.otherMemos = [];
    kvSet('remoteTarget', null);
    S.myMemos = await memos.getMyMemos(session.id);
    await saveSettings({ lastProjectId: projectId, lastTemplateId: templateId, lastScene: scene });
    S.setupTitle = ''; S.setupPeople = ''; S.setupConf = false; S.showLocalRec = false;
    render();
  } catch (e) {
    toast(errorMessage(e), 'error', 8000);
  } finally {
    if (btn) btn.disabled = false;
  }
}

const renderHeaderThrottled = debounce(renderHeader, 200);

function updateMeter() {
  const f = $('#meter-fill');
  if (!f) return;
  const pct = Math.max(0, Math.min(100, ((S.level + 60) / 60) * 100));
  f.style.width = `${pct}%`;
  f.classList.toggle('low', S.level < -50);
}

async function startOnline(session) {
  try {
    await pipe.ensureFolder(session);
    await pipe.writeMeta(session);
    S.publisher = new LivePublisher({
      session,
      getElapsed: () => (S.rec ? S.rec.elapsed : 0),
      onRemoteStop: (c) => { toast(`${c.by || '他の端末'}から終了の指示を受けました`, 'ok'); handleStop(true); },
    });
    await S.publisher.start();
    await memos.flushMemos(session.id, session.folderPath).catch(() => {});
    if (S.view === 'meeting') renderRec();
  } catch (e) {
    toast(`OneDriveとの同期を開始できませんでした。録音は続いています（${errorMessage(e)}）`, 'warn', 7000);
  }
}

async function handleStop() {
  const rec = S.rec;
  if (!rec) return;
  S.rec = null;
  const s = S.session;
  try {
    const res = await rec.stop();
    s.startedAt = s.startedAt || res.startedAt || s.createdAt;
    s.durationSec = res.durationSec;
    s.levels = res.levels;
    s.gaps = res.gaps;
    s.endedAt = s.startedAt + res.durationSec * 1000;
    s.status = 'recorded';
    await putSession(s);
    if (S.publisher) { await S.publisher.stop('processing'); S.publisher = null; }
    toast(`録音を終了しました（${fmtTime(res.durationSec)}）`, 'ok');
  } catch (e) {
    toast(errorMessage(e), 'error');
  }
  render();
  if (getSettings().recording.autoStep1) runStep1UI(s);
}

async function withWakeLock(fn) {
  let wake = null;
  try { if ('wakeLock' in navigator) wake = await navigator.wakeLock.request('screen'); } catch { /* 無視 */ }
  try { return await fn(); } finally { try { if (wake) await wake.release(); } catch { /* 無視 */ } }
}

async function runStep1UI(s) {
  if (S.proc) { toast('別の処理が実行中です', 'warn'); return; }
  if (!getSettings().onedrive.clientId) { toast('OneDriveを設定すると、保存と書き起こしを実行できます（録音は端末内に保存済み）', 'warn', 7000); return; }
  try { if (!(await ensureMs(false))) { toast('OneDriveに再サインインしてから「再試行」を押してください', 'warn', 7000); render(); return; } } catch (e) { toast(errorMessage(e), 'error'); return; }
  S.proc = { sessionId: s.id, kind: 'step1', progress: { text: '開始しています' } };
  renderProc();
  await withWakeLock(async () => {
    try {
      await pipe.runStep1(s, { onProgress: (p) => { if (p && S.proc) { S.proc.progress = p; renderProc(); } } });
      toast('書き起こしが完了しました。続けて「② 議事録作成」を押せます。', 'ok', 7000);
    } catch (e) {
      toast(`①で問題が起きました：${errorMessage(e)}`, 'error', 9000);
    } finally {
      S.proc = null;
      render();
    }
  });
}

async function runStep2UI(target) {
  if (S.proc) { toast('別の処理が実行中です', 'warn'); return; }
  const st = getSettings();
  try {
    // ボタンを押した直後でないと認証画面がブロックされるため、通信を伴うOneDriveの確認より先に行う
    if (st.google.clientId && st.google.autoCalendar && !(await googleTokenValid())) {
      try { await googleLogin(st.google.clientId, st.google.account); S.status.google = true; } catch { /* カレンダーなしで続行 */ }
    }
    if (!(await ensureMs(true))) { toast('OneDriveの設定が必要です', 'warn'); return; }
  } catch (e) { toast(errorMessage(e), 'error'); return; }
  S.proc = { sessionId: target.id, kind: 'step2', progress: { text: '準備しています' } };
  renderProc();
  await withWakeLock(async () => {
    try {
      let s = await getSession(target.id);
      if ((!s || !s.folderPath || !(s.steps && s.steps.transcript)) && target.folderPath) s = await pipe.loadRemoteSession(target.folderPath);
      if (!s) throw new Error('会議が見つかりません');
      await pipe.runStep2(s, { onProgress: (p) => { if (p && S.proc) { S.proc.progress = p; renderProc(); } } });
      if (S.session && S.session.id === s.id) S.session = s;
      if (S.remote && S.remote.sessionId === s.id) { S.remoteMeta = pipe.sessionMeta(s); }
      const learned = s.learnedTerms ? `（資料から${s.learnedTerms}語を単語帳に追加）` : '';
      toast(`議事録を作成し、OneDriveに保存しました${learned}`, 'ok', 7000);
      S.proc = null;
      openDetail(s, 'minutes');
      return;
    } catch (e) {
      toast(`②で問題が起きました：${errorMessage(e)}`, 'error', 9000);
    } finally {
      S.proc = null;
      if (S.view !== 'detail') render(); else renderDetailBody();
    }
  });
}

async function saveMemosUI(target) {
  try {
    if (!(await ensureMs(true))) return;
    await pipe.saveMemosOnly(target);
    toast('メモをOneDriveに保存しました', 'ok');
  } catch (e) { toast(errorMessage(e), 'error'); }
}

function renderProc() {
  const box = $('#proc-box');
  if (!box) { if (S.view === 'detail') renderDetailBody(); return; }
  clear(box);
  if (S.proc) {
    const p = S.proc.progress || {};
    box.appendChild(el('section', { class: 'card' },
      el('h2', {}, S.proc.kind === 'step1' ? '① 保存・書き起こし中' : '② 議事録を作成中'),
      el('p', {}, p.text || ''),
      progressBar(p.ratio),
      el('p', { class: 'muted small' }, '完了まで画面をつけたままにしてください。')));
    return;
  }
  const s = S.session;
  if (!s || S.rec) return;
  const card = el('section', { class: 'card' }, el('h2', {}, `直前の会議：${s.title}`),
    el('p', {}, `${fmtDateTime(s.startedAt)}　${fmtTime(s.durationSec)}　${statusLabel(s)}`),
    s.error ? el('p', { class: 'warn small' }, s.error) : null,
    s.calendarPending ? el('p', { class: 'muted small' }, 'カレンダーは、②の実行時か「カレンダー登録」で登録します。') : null);
  const row = el('div', { class: 'row' });
  if (!(s.steps && s.steps.transcript)) row.appendChild(el('button', { class: 'btn primary', type: 'button', onclick: () => runStep1UI(s) }, s.status === 'error' ? '① を再試行' : '① 保存・書き起こし'));
  else row.appendChild(el('button', { class: 'btn primary', type: 'button', onclick: () => runStep2UI(s) }, s.steps.minutes ? '② 議事録を作り直す' : '② メモを組み込んで議事録作成'));
  row.appendChild(el('button', { class: 'btn', type: 'button', onclick: () => saveMemosUI(s) }, '📌 メモだけ保存'));
  row.appendChild(el('button', { class: 'btn ghost', type: 'button', onclick: () => openDetail(s) }, '詳細'));
  row.appendChild(el('button', { class: 'btn ghost', type: 'button', onclick: () => { S.session = null; render(); } }, '閉じる'));
  card.appendChild(row);
  box.appendChild(card);
}

// ---- メモ欄（B-5 / B-6） ----
function renderMemoBox() {
  const box = $('#memo-box');
  if (!box) return;
  clear(box);
  const target = memoTarget();
  const st = getSettings();
  const where = S.rec ? 'この端末の録音' : S.remote ? `${S.remote.deviceName}の録音「${S.remote.title}」` : target.isDraft ? '次に開始する録音' : `「${target.title}」`;
  const prev = $('#memo-input');
  const hadFocus = prev && document.activeElement === prev;
  const input = el('textarea', {
    class: 'input memo-input', id: 'memo-input', rows: isMobile() ? 3 : 5,
    placeholder: 'メモを入力（Enterで確定、Shift+Enterで改行、空欄でSpaceを押すと★ピン留め）',
  });
  input.value = S.memoDraftText || '';
  input.addEventListener('input', () => {
    S.memoDraftText = input.value;
    if (input.value.trim() && S.memoT0 == null) S.memoT0 = Date.now();
    if (!input.value.trim()) S.memoT0 = null;
    const lock = $('#memo-lock');
    if (lock) lock.textContent = S.memoT0 && Number.isFinite(target.startedAt) ? `この位置に記録：${fmtTime((S.memoT0 - target.startedAt) / 1000)}` : '';
  });
  input.addEventListener('keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return; // 日本語入力の変換中は無視
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); submitMemo(); }
    else if (e.key === ' ' && !input.value) { e.preventDefault(); quickMemo('star'); }
  });
  box.appendChild(el('section', { class: 'card memo-card' },
    el('h2', {}, 'メモ'),
    el('p', { class: 'muted small' }, `記録先：${where}`),
    input,
    el('div', { class: 'row between' },
      el('span', { id: 'memo-lock', class: 'muted small mono' }),
      el('div', { class: 'row' },
        el('button', { class: 'btn', type: 'button', onclick: () => quickMemo('star') }, '★'),
        el('button', { class: 'btn primary', type: 'button', onclick: submitMemo }, 'メモ追加'))),
    el('div', { class: 'phrases' }, (st.memo.phrases || []).map((p) => el('button', { class: 'btn small ghost', type: 'button', onclick: () => quickMemo('phrase', p) }, p))),
    el('div', { id: 'memo-list', class: 'memo-list' })));
  if (hadFocus) input.focus();
  loadMemosFor(target).then(renderMemoList);
}

async function loadMemosFor(target) {
  S.myMemos = await memos.getMyMemos(target.id);
}

function currentMemos() {
  const target = memoTarget();
  const all = memos.combine(S.myMemos, [S.otherMemos]);
  return all.map((m) => ({ ...m, offset: tp.memoOffset(m, target.startedAt) }));
}

function renderMemoList() {
  const box = $('#memo-list');
  if (!box) return;
  clear(box);
  const list = currentMemos();
  const mine = getSettings().deviceId;
  if (!list.length) { box.appendChild(el('p', { class: 'muted small' }, 'まだメモはありません。')); return; }
  for (const m of list.slice().reverse()) {
    const time = Number.isFinite(m.offset) ? (m.offset < 0 ? '開始前' : fmtTime(m.offset)) : '開始前';
    const own = m.deviceId === mine;
    box.appendChild(el('div', { class: `memo ${m.kind === 'star' ? 'star' : ''}` },
      el('span', { class: 'mono time' }, time),
      el('span', { class: 'memo-text' }, m.kind === 'star' && !m.text.startsWith('★') ? `★${m.text}` : m.text),
      el('span', { class: 'muted small' }, m.deviceName || ''),
      own ? el('button', { class: 'icon-btn', type: 'button', title: '編集', onclick: () => editMemo(m) }, '✎') : null,
      own ? el('button', { class: 'icon-btn', type: 'button', title: '削除', onclick: () => removeMemo(m) }, '×') : null));
  }
}

async function submitMemo() {
  const input = $('#memo-input');
  if (!input) return;
  const text = input.value.trim();
  if (!text) return;
  const t0 = S.memoT0 || Date.now();
  // 先に入力欄を空にする（続けてすぐ次のメモを書けるように）
  input.value = '';
  S.memoDraftText = '';
  S.memoT0 = null;
  const lock = $('#memo-lock');
  if (lock) lock.textContent = '';
  const target = memoTarget();
  await memos.addMemo(target, { text, t0 });
  await loadMemosFor(target);
  renderMemoList();
  input.focus();
}

async function quickMemo(kind, text = '') {
  const target = memoTarget();
  await memos.addMemo(target, { text: kind === 'star' ? '★要確認' : text, kind, t0: Date.now() });
  await loadMemosFor(target);
  renderMemoList();
  toast(kind === 'star' ? '★ピン留めしました' : `「${text}」を記録しました`, 'ok', 1500);
}

async function editMemo(m) {
  const r = await formDialog('メモを編集', [{ name: 'text', label: 'メモ', value: m.text }]);
  if (!r) return;
  const target = memoTarget();
  await memos.updateMemo(target, m.id, { text: r.text });
  await loadMemosFor(target);
  renderMemoList();
}

async function removeMemo(m) {
  if (!(await confirmDialog('このメモを削除しますか？', { ok: '削除', danger: true }))) return;
  const target = memoTarget();
  await memos.deleteMemo(target, m.id);
  await loadMemosFor(target);
  renderMemoList();
}

async function refreshRemoteMemos() {
  const target = S.view === 'detail' && S.detail ? S.detail : memoTarget();
  if (!target || !target.folderPath || !S.status.ms || document.visibilityState !== 'visible') return;
  try {
    const lists = await readAllMemos(target.folderPath);
    const mine = getSettings().deviceId;
    S.otherMemos = lists.flat().filter((m) => m.deviceId !== mine);
    if (S.view === 'meeting') renderMemoList();
  } catch { /* 次回 */ }
}

// ===========================================================================
// 一覧（F-18）
// ===========================================================================
function viewList() {
  const q = el('input', { class: 'input', placeholder: '会議名で絞り込み', value: S.listQuery, oninput: (e) => { S.listQuery = e.target.value; renderListItems(); } });
  const searchBtn = el('button', { class: 'btn', type: 'button', onclick: () => searchOneDrive(q.value) }, 'OneDriveの中身を検索');
  return el('div', { class: 'col' },
    el('section', { class: 'card' }, el('h2', {}, '会議一覧'),
      el('div', { class: 'row nowrap' }, q, searchBtn),
      el('div', { id: 'search-hits' }),
      el('div', { id: 'list-items', class: 'list' }, el('p', { class: 'muted' }, '読み込み中…'))));
}

async function loadList() {
  try {
    if (S.status.ms) await ensureMs(false).catch(() => {});
    S.list = await pipe.listAll({ remote: S.status.ms });
  } catch (e) {
    S.list = await pipe.listAll({ remote: false });
    toast(errorMessage(e), 'warn');
  }
  renderListItems();
}

function renderListItems() {
  const box = $('#list-items');
  if (!box) return;
  clear(box);
  const q = S.listQuery.trim();
  const items = S.list.filter((s) => !q || (s.title || '').includes(q) || (s.projectName || '').includes(q));
  if (!items.length) { box.appendChild(el('p', { class: 'muted' }, '会議はまだありません。')); return; }
  for (const s of items) {
    const proj = vocab.projectById(s.projectId);
    box.appendChild(el('button', { class: 'list-item', type: 'button', onclick: () => openListItem(s) },
      el('div', { class: 'row between' }, el('strong', {}, s.title || '（無題）'), chip(statusLabel(s), s.status === 'done' ? 'ok' : s.status === 'error' ? 'warn' : 'muted')),
      el('span', { class: 'muted small' }, `${fmtDateTime(s.startedAt)}　${fmtTime(s.durationSec || 0)}　${proj && proj.id !== 'common' ? proj.name : ''}${s.confidential ? '　🔒機密' : ''}${s.where === 'local' ? '　（この端末のみ）' : ''}`)));
  }
}

async function openListItem(s) {
  try {
    const local = await getSession(s.id);
    if (local && (local.transcript || !s.folderPath)) { openDetail(local); return; }
    await openDetailByFolder(s.folderPath);
  } catch (e) { toast(errorMessage(e), 'error'); }
}

async function searchOneDrive(q) {
  const box = $('#search-hits');
  clear(box);
  if (!q.trim()) return;
  try {
    if (!(await ensureMs(true))) return;
    box.appendChild(el('p', { class: 'muted small' }, '検索中…'));
    const hits = await od.search(rootPath(), q.trim());
    clear(box);
    const folders = new Map();
    for (const h of hits) {
      const parentPath = h.parentReference && h.parentReference.path ? decodeURIComponent(h.parentReference.path.replace(/^\/drive\/root:\/?/, '')) : '';
      const m = parentPath.match(/(\d{8}_\d{4}(?:_\d+)?)/);
      if (!m) continue;
      const folderPath = parentPath.slice(0, parentPath.indexOf(m[1]) + m[1].length);
      if (!folders.has(folderPath)) folders.set(folderPath, []);
      folders.get(folderPath).push(h.name);
    }
    if (!folders.size) { box.appendChild(el('p', { class: 'muted small' }, '見つかりませんでした（OneDriveの検索に反映されるまで数分かかることがあります）。')); return; }
    box.appendChild(el('h3', {}, `「${q}」を含む会議`));
    for (const [fp, names] of folders) {
      box.appendChild(el('button', { class: 'list-item', type: 'button', onclick: () => openDetailByFolder(fp) },
        el('strong', {}, fp.split('/').pop()), el('span', { class: 'muted small' }, names.join('、'))));
    }
  } catch (e) { clear(box); toast(errorMessage(e), 'error'); }
}

// ===========================================================================
// 会議の詳細（再生・書き起こし・メモ・議事録・資料）
// ===========================================================================
async function openDetailByFolder(folderPath) {
  try {
    if (!(await ensureMs(true))) return;
    toast('読み込んでいます…', 'info', 1500);
    const s = await pipe.loadRemoteSession(folderPath);
    openDetail(s);
  } catch (e) { toast(errorMessage(e), 'error'); }
}

async function openDetail(s, tab = 'transcript') {
  leaveDetail();
  S.detail = s;
  S.detailTab = tab;
  S.editTranscript = false;
  S.editMinutes = false;
  S.myMemos = await memos.getMyMemos(s.id);
  S.otherMemos = [];
  S.view = 'detail';
  render();
  window.scrollTo(0, 0);
  refreshRemoteMemos().then(() => { if (S.view === 'detail' && S.detailTab === 'memos') renderDetailBody(); });
  // 音声：端末内にあればそれを、なければOneDriveから
  try {
    const chunks = await getChunks(s.id);
    if (chunks.length) S.audioUrl = URL.createObjectURL(wavFromPcmBlob(pcmBlobFromChunks(chunks)));
    else if (s.folderPath && S.status.ms) S.audioUrl = await od.downloadUrl(od.joinPath(s.folderPath, '録音データ.wav'));
    const a = $('#player');
    if (a && S.audioUrl) a.src = S.audioUrl;
  } catch { /* 音声なしで表示 */ }
}

function leaveDetail() {
  if (S.audioUrl && S.audioUrl.startsWith('blob:')) URL.revokeObjectURL(S.audioUrl);
  S.audioUrl = null;
  S.detail = null;
}

function seek(sec) {
  const a = $('#player');
  if (!a) return;
  a.currentTime = Math.max(0, sec - 1);
  a.play().catch(() => {});
}

function viewDetail() {
  const s = S.detail;
  const wrap = el('div', { class: 'col' });
  if (!s) return wrap;
  const player = el('audio', { id: 'player', controls: true, preload: 'metadata' });
  if (S.audioUrl) player.src = S.audioUrl;
  const skip = el('input', { type: 'checkbox', id: 'skip-silence' });
  const speed = el('select', { class: 'input small', onchange: (e) => { player.playbackRate = Number(e.target.value); } },
    ['1', '1.25', '1.5', '2'].map((v) => el('option', { value: v }, `${v}倍`)));
  player.addEventListener('timeupdate', () => {
    const t = Math.floor(player.currentTime);
    const lv = s.levels || [];
    if (skip.checked && lv.length && lv[t] < -48 && lv[t + 1] < -48) {
      let j = t + 1;
      while (j < lv.length && lv[j] < -48) j++;
      if (j - t >= 2) player.currentTime = Math.max(0, j - 0.3);
    }
    highlightSegment(player.currentTime);
  });
  const st = getSettings();
  const actions = el('div', { class: 'row' },
    !(s.steps && s.steps.transcript) ? el('button', { class: 'btn primary', type: 'button', disabled: !!S.proc, onclick: () => runStep1UI(s) }, '① 保存・書き起こし') : null,
    s.steps && s.steps.transcript ? el('button', { class: 'btn primary', type: 'button', disabled: !!S.proc, onclick: () => runStep2UI(s) }, s.minutes ? '② 議事録を作り直す' : '② 議事録作成') : null,
    s.transcript ? el('button', { class: 'btn', type: 'button', onclick: () => renameSpeakersUI(s) }, '話者名を変更') : null,
    el('button', { class: 'btn', type: 'button', onclick: () => calendarUI(s) }, '📅 カレンダー登録'),
    s.folderWebUrl ? el('a', { class: 'btn ghost', href: s.folderWebUrl, target: '_blank', rel: 'noopener noreferrer' }, 'OneDriveで開く') : null,
    el('button', { class: 'btn ghost', type: 'button', onclick: () => go('list') }, '一覧へ戻る'));
  wrap.append(
    el('section', { class: 'card' },
      el('h2', {}, s.title || '（無題）'),
      el('p', { class: 'muted' }, `${fmtDateTime(s.startedAt)}　${fmtTime(s.durationSec || 0)}　${statusLabel(s)}${s.confidential ? '　🔒機密' : ''}${s.deviceName ? `　録音：${s.deviceName}` : ''}`),
      s.error ? el('p', { class: 'warn small' }, s.error) : null,
      player,
      el('div', { class: 'row' }, el('label', { class: 'check' }, skip, '無音を飛ばす'), speed,
        el('button', { class: 'btn small ghost', type: 'button', onclick: () => jumpStar(1) }, '次の★へ'),
        el('button', { class: 'btn small ghost', type: 'button', onclick: () => jumpStar(-1) }, '前の★へ')),
      actions,
      st.google.clientId ? null : (s.calendarUrl ? el('p', { class: 'muted small' }, 'カレンダー連携が未設定のため、「カレンダー登録」で登録画面を開きます。') : null)),
    el('div', { class: 'tabs' }, [['transcript', '書き起こし'], ['memos', 'メモ'], ['minutes', '議事録'], ['materials', '資料']].map(([id, label]) =>
      el('button', { type: 'button', class: `tab ${S.detailTab === id ? 'active' : ''}`, onclick: () => { S.detailTab = id; render(); } }, label))),
    el('div', { id: 'detail-body' }));
  queueMicrotask(renderDetailBody);
  return wrap;
}

function jumpStar(dir) {
  const a = $('#player');
  if (!a) return;
  const stars = currentDetailMemos().filter((m) => m.kind === 'star' && Number.isFinite(m.offset)).map((m) => m.offset).sort((x, y) => x - y);
  const now = a.currentTime;
  const next = dir > 0 ? stars.find((t) => t > now + 2) : stars.reverse().find((t) => t < now - 3);
  if (next == null) { toast('★はありません'); return; }
  seek(next);
}

function currentDetailMemos() {
  const s = S.detail;
  return memos.combine(S.myMemos, [S.otherMemos]).map((m) => ({ ...m, offset: tp.memoOffset(m, s.startedAt) }));
}

let lastHi = -1;
function highlightSegment(t) {
  if (S.detailTab !== 'transcript' || !S.detail || !S.detail.transcript) return;
  const segs = S.detail.transcript.segments;
  let idx = -1;
  for (let i = 0; i < segs.length; i++) { if (segs[i].start <= t + 0.5) idx = i; else break; }
  if (idx === lastHi) return;
  const prev = document.querySelector('.seg.playing');
  if (prev) prev.classList.remove('playing');
  const cur = document.querySelector(`.seg[data-i="${idx}"]`);
  if (cur) cur.classList.add('playing');
  lastHi = idx;
}

function renderDetailBody() {
  const box = $('#detail-body');
  const s = S.detail;
  if (!box || !s) return;
  clear(box);
  if (S.proc && S.proc.sessionId === s.id) {
    const p = S.proc.progress || {};
    box.appendChild(el('section', { class: 'card' }, el('p', {}, p.text || '処理中…'), progressBar(p.ratio)));
  }
  if (S.detailTab === 'transcript') box.appendChild(detailTranscript(s));
  else if (S.detailTab === 'memos') box.appendChild(detailMemos(s));
  else if (S.detailTab === 'minutes') box.appendChild(detailMinutes(s));
  else box.appendChild(detailMaterials(s));
}

function detailTranscript(s) {
  const card = el('section', { class: 'card' });
  if (!s.transcript || !s.transcript.segments) {
    card.appendChild(el('p', { class: 'muted' }, '書き起こしはまだありません。「① 保存・書き起こし」を実行してください。'));
    return card;
  }
  const segs = s.transcript.segments;
  const mm = currentDetailMemos();
  const edits = {};
  card.appendChild(el('div', { class: 'row between' },
    el('span', { class: 'muted small' }, `発言 ${segs.length}件${s.transcript.corrections ? `／用語の補正 ${s.transcript.corrections}件` : ''}${s.refineError ? '（用語補正は失敗）' : ''}`),
    S.editTranscript
      ? el('div', { class: 'row' },
        el('button', { class: 'btn ghost small', type: 'button', onclick: () => { S.editTranscript = false; renderDetailBody(); } }, 'やめる'),
        el('button', {
          class: 'btn primary small', type: 'button',
          onclick: async () => {
            try {
              const r = await pipe.saveTranscriptEdits(s, edits);
              S.editTranscript = false;
              toast(`${r.changed}件の手直しを保存しました${r.learned ? `。単語帳に${r.learned}語を追加しました` : ''}`, 'ok', 6000);
              renderDetailBody();
            } catch (e) { toast(errorMessage(e), 'error'); }
          },
        }, '保存して学習'))
      : el('button', { class: 'btn small', type: 'button', onclick: () => { S.editTranscript = true; renderDetailBody(); } }, '手直しする')));
  const list = el('div', { class: 'segs' });
  segs.forEach((seg, i) => {
    const near = tp.memosNear(seg, mm, 12).filter((m) => !m._shown);
    near.forEach((m) => { m._shown = true; });
    const body = S.editTranscript
      ? el('textarea', { class: 'input seg-edit', rows: Math.min(6, Math.ceil(seg.text.length / 40) + 1), oninput: (e) => { edits[seg.id] = e.target.value; } }, seg.text)
      : el('span', { class: 'seg-text' }, seg.text);
    list.appendChild(el('div', { class: 'seg', dataset: { i: String(i) } },
      el('button', { class: 'time-tag', type: 'button', onclick: () => seek(seg.start) }, `[${fmtTime(seg.start)}]`),
      el('strong', { class: 'spk' }, `${seg.speaker}：`),
      body,
      seg.estimated ? el('span', { class: 'muted small', title: '時刻は推定です' }, '〜') : null,
      near.length ? el('div', { class: 'seg-memos' }, near.map((m) => el('div', { class: 'memo inline' }, el('span', { class: 'mono time' }, fmtTime(m.offset)), m.text))) : null));
  });
  card.appendChild(list);
  return card;
}

function detailMemos(s) {
  const card = el('section', { class: 'card' });
  const input = el('input', { class: 'input', placeholder: '再生中の位置にメモを追加' });
  const add = async () => {
    const text = input.value.trim();
    if (!text) return;
    const a = $('#player');
    await memos.addMemo(s, { text, offset: a ? Math.round(a.currentTime * 10) / 10 : null });
    S.myMemos = await memos.getMyMemos(s.id);
    input.value = '';
    renderDetailBody();
  };
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.isComposing && e.keyCode !== 229) add(); });
  card.appendChild(el('div', { class: 'row nowrap' }, input, el('button', { class: 'btn primary', type: 'button', onclick: add }, '追加')));
  const list = currentDetailMemos();
  const mine = getSettings().deviceId;
  if (!list.length) card.appendChild(el('p', { class: 'muted' }, 'メモはありません。'));
  for (const m of list) {
    card.appendChild(el('div', { class: `memo ${m.kind === 'star' ? 'star' : ''}` },
      el('button', { class: 'time-tag', type: 'button', onclick: () => seek(m.offset || 0) }, Number.isFinite(m.offset) ? `[${fmtTime(Math.max(0, m.offset))}]` : '[--:--]'),
      el('span', { class: 'memo-text' }, m.text),
      el('span', { class: 'muted small' }, m.deviceName || ''),
      m.deviceId === mine ? el('button', {
        class: 'icon-btn', type: 'button', title: '削除',
        onclick: async () => { await memos.deleteMemo(s, m.id); S.myMemos = await memos.getMyMemos(s.id); renderDetailBody(); },
      }, '×') : null));
  }
  card.appendChild(el('div', { class: 'row' }, el('button', { class: 'btn', type: 'button', onclick: () => saveMemosUI(s) }, '📌 メモをOneDriveに保存')));
  return card;
}

function detailMinutes(s) {
  const card = el('section', { class: 'card' });
  if (!s.minutes || !s.minutes.markdown) {
    card.appendChild(el('p', { class: 'muted' }, '議事録はまだありません。'));
    if (s.steps && s.steps.transcript) card.appendChild(el('button', { class: 'btn primary', type: 'button', disabled: !!S.proc, onclick: () => runStep2UI(s) }, '② メモを組み込んで議事録作成'));
    return card;
  }
  if (S.editMinutes) {
    const ta = el('textarea', { class: 'input minutes-edit', rows: 24 }, s.minutes.markdown);
    card.append(
      el('p', { class: 'muted small' }, '直した内容から、用語と書き方のルールを学び、次回以降に反映します。'),
      ta,
      el('div', { class: 'row end' },
        el('button', { class: 'btn ghost', type: 'button', onclick: () => { S.editMinutes = false; renderDetailBody(); } }, 'やめる'),
        el('button', {
          class: 'btn primary', type: 'button',
          onclick: async () => {
            try {
              const r = await pipe.saveMinutesEdits(s, ta.value);
              S.editMinutes = false;
              toast(`議事録を保存しました${r.learnedTerms || r.learnedRules ? `（用語${r.learnedTerms}件・ルール${r.learnedRules}件を学習）` : ''}`, 'ok', 6000);
              renderDetailBody();
            } catch (e) { toast(errorMessage(e), 'error'); }
          },
        }, '保存して学習')));
    return card;
  }
  card.append(
    el('div', { class: 'row between' },
      el('span', { class: 'muted small' }, s.minutes.flagged || s.minutes.weak
        ? `⚠根拠未確認 ${s.minutes.flagged || 0}件、△内容要確認 ${s.minutes.weak || 0}件。[MM:SS]を押して発言を確かめてください。`
        : '[MM:SS]を押すと、根拠の発言を再生します。'),
      el('div', { class: 'row' },
        el('button', { class: 'btn small', type: 'button', onclick: () => copyText(s.minutes.markdown) }, 'コピー'),
        el('button', { class: 'btn small', type: 'button', onclick: () => { S.editMinutes = true; renderDetailBody(); } }, '手直しする'))),
    renderMarkdown(s.minutes.markdown, seek));
  return card;
}

/** OneDriveにサインインして、保存先フォルダを用意する（利用者の操作の直後に呼ぶこと） */
async function signInOneDrive() {
  let step = 'signin';
  try {
    await msLogin(getSettings().onedrive.clientId);
    S.status.ms = true;
    step = 'folder';
    await od.ensureFolder(rootPath());
    const me = await od.getMe().catch(() => ({}));
    S.msAccount = me.userPrincipalName || me.mail || me.displayName || '';
    toast(`OneDriveにサインインし、保存先フォルダ「${rootPath()}」を用意しました${S.msAccount ? `（${S.msAccount}）` : ''}`, 'ok', 8000);
    afterMsReady();
    render();
    return true;
  } catch (e) {
    toast(step === 'folder' ? `サインインはできましたが、保存先フォルダ「${rootPath()}」を作れませんでした：${errorMessage(e)}` : errorMessage(e), 'error', 15000);
    renderHeader();
    return false;
  }
}

// ---------------------------------------------------------------------------
// ほかの端末への設定の受け渡し（QRコード）。サーバーを通さず、画面からカメラへ直接渡す
// ---------------------------------------------------------------------------
function showSetupQr(includeKey) {
  const st = getSettings();
  if (!st.onedrive.clientId) { toast('先にこの端末で、OneDriveのアプリIDを設定してください。', 'warn'); return; }
  if (includeKey && !st.gemini.apiKey) toast('この端末にはGeminiのAPIキーが保存されていないため、キーは渡しません。', 'warn', 6000);
  let canvas;
  try { canvas = qrCanvas(`${new URL('./', document.baseURI).href}${SETUP_PREFIX}${encodeSetup(setupPayload(getSettings(), includeKey))}`, { scale: 5 }); } catch (e) { toast(errorMessage(e), 'error'); return; }
  canvas.className = 'qr';
  const keyIncluded = includeKey && !!st.gemini.apiKey;
  const dlg = el('dialog', { class: 'dialog' },
    el('h3', {}, 'ほかの端末を設定する'),
    el('ol', { class: 'small' },
      el('li', {}, 'スマホのカメラ（またはGoogleレンズ）で、このQRコードを読み取ります。'),
      el('li', {}, '表示されたリンクを開き、アプリの画面で「取り込む」を押します。'),
      el('li', {}, '続けて「OneDriveにサインイン」を押します。これで完了です。')),
    canvas,
    keyIncluded ? el('p', { class: 'warn small' }, 'このQRコードにはGeminiのAPIキーが含まれます。読み取ったらすぐ閉じ、ほかの人に見せたり、写真に撮って残したりしないでください。') : null,
    el('div', { class: 'row end' }, el('button', { class: 'btn primary', type: 'button', onclick: () => dlg.close() }, '閉じる')));
  dlg.addEventListener('close', () => dlg.remove());
  document.body.appendChild(dlg);
  dlg.showModal();
}

/** 起動時に #setup= で開かれた場合：内容を確かめてもらってから取り込む */
async function offerSetupImport(raw) {
  const p = parseSetup(raw);
  if (!p) { toast('設定のQRコードを読み取れませんでした。もう一度表示し直して読み取ってください。', 'error', 8000); return; }
  const st = getSettings();
  const tail = (v) => `…${String(v).slice(-4)}`;
  const overwrite = (st.onedrive.clientId && st.onedrive.clientId !== p.od) || (p.g && st.gemini.apiKey && st.gemini.apiKey !== p.g);
  const ok = await new Promise((resolve) => {
    const dlg = el('dialog', { class: 'dialog' },
      el('h3', {}, 'ほかの端末の設定を取り込みます'),
      el('ul', { class: 'small' },
        el('li', {}, `OneDrive：アプリID ${tail(p.od)}、保存先「${p.r || '会議録音アーカイブ'}」`),
        el('li', {}, p.g ? `Gemini：APIキー ${tail(p.g)}${p.p ? '（課金を有効にしたキー）' : ''}` : 'Gemini：キーは含まれていません（この端末で貼り付けます）'),
        el('li', {}, p.gc ? `Googleカレンダー：${p.ga || 'アカウント未指定'}` : 'Googleカレンダー：設定なし')),
      el('p', { class: 'warn small' }, 'ご自身の端末に表示したQRコードを読み取った場合だけ「取り込む」を押してください。人から送られてきたリンクの場合は「やめる」を押してください（会議の内容の送り先を書き換えられるおそれがあります）。'),
      overwrite ? el('p', { class: 'warn small strong' }, 'この端末の今の設定を上書きします。') : null,
      el('div', { class: 'row end' },
        el('button', { class: 'btn ghost', type: 'button', onclick: () => { dlg.close(); resolve(false); } }, 'やめる'),
        el('button', { class: 'btn primary', type: 'button', onclick: () => { dlg.close(); resolve(true); } }, '取り込む')));
    dlg.addEventListener('close', () => { dlg.remove(); resolve(false); });
    document.body.appendChild(dlg);
    dlg.showModal();
  });
  if (!ok) { toast('取り込みをやめました。', 'info'); return; }
  const patch = { onedrive: { clientId: p.od, rootPath: p.r || '会議録音アーカイブ' } };
  if (p.g) patch.gemini = { apiKey: p.g, paidConfirmed: !!p.p };
  if (p.gc) patch.google = { clientId: p.gc, account: p.ga || '' };
  await saveSettings(patch);
  renderHeader();
  render();
  // サインインの画面は、利用者が押した直後でないと開けないため、ボタンを押してもらう
  const dlg = el('dialog', { class: 'dialog' },
    el('h3', {}, '取り込みました'),
    el('p', { class: 'small' }, '最後に、この端末でOneDriveにサインインしてください。Microsoftの決まりで、サインインだけは端末ごとに1回必要です。'),
    p.gc ? el('p', { class: 'muted small' }, 'Googleカレンダーは、この端末で②（議事録作成）を実行するときに、必要なら確認の画面が出ます。') : null,
    el('div', { class: 'row end' },
      el('button', { class: 'btn ghost', type: 'button', onclick: () => dlg.close() }, 'あとで'),
      el('button', { class: 'btn primary', type: 'button', onclick: async () => { dlg.close(); await signInOneDrive(); } }, 'OneDriveにサインイン')));
  dlg.addEventListener('close', () => dlg.remove());
  document.body.appendChild(dlg);
  dlg.showModal();
}

/** 保存先フォルダの入力を、OneDriveの一番上からのパスにそろえる（「OneDrive/」や前後の「/」を取り除く） */
function normalizeRoot(v) {
  const p = String(v || '').trim().replace(/\\/g, '/').replace(/^\/+|\/+$/g, '').replace(/^OneDrive\/+/i, '').replace(/\/{2,}/g, '/');
  return p || '会議録音アーカイブ';
}

/** サインインの状態と保存先フォルダを表示し、OneDriveで直接開けるようにする */
function renderOdStatus(box) {
  box.replaceChildren();
  if (!S.status.ms) {
    box.appendChild(el('p', { class: 'muted small' }, '未サインイン：保存先フォルダはまだ作られていません。'));
    return;
  }
  const line = el('p', { class: 'small' }, `サインイン中${S.msAccount ? `：${S.msAccount}` : ''}　保存先：OneDrive / ${rootPath()}`);
  const open = el('button', { class: 'btn', type: 'button', onclick: async () => {
    const w = window.open('', '_blank');
    try {
      const url = await od.webUrl(rootPath());
      if (!url) { if (w) w.close(); toast(`OneDriveに「${rootPath()}」が見つかりません。「保存して再サインイン」を押すと作り直します。`, 'warn', 8000); return; }
      if (w) w.location.href = url; else window.open(url, '_blank');
    } catch (e) { if (w) w.close(); toast(errorMessage(e), 'error'); }
  } }, '保存先をOneDriveで開く');
  box.append(line, el('div', { class: 'row' }, open));
}

/** サインインの戻り先URL。各サービスの登録画面と1文字でも違うとサインインできないため、コピーして貼れるようにする */
function redirectBox() {
  const uri = redirectUri();
  const box = el('div', { class: 'redirect-box' },
    el('p', { class: 'muted small' }, 'このアプリのサインイン戻り先（Microsoft Entraの「認証」→「シングルページ アプリケーション」と、Googleの「承認済みのリダイレクトURI」に、このとおり登録します）'),
    el('div', { class: 'row' },
      el('input', { class: 'input mono', value: uri, readonly: true, onfocus: (e) => e.target.select() }),
      el('button', { class: 'btn', type: 'button', onclick: () => copyText(uri) }, 'コピー')));
  if (!/^https:\/\//.test(uri) && !/^http:\/\/(localhost|127\.0\.0\.1)(:|\/)/.test(uri)) {
    box.appendChild(el('p', { class: 'warn small' }, 'このページは公開先（https://…github.io/…）から開かれていません。この状態ではサインインできません。GitHub PagesのURLから開き直してください。'));
  }
  return box;
}

async function copyText(text) {
  try { await navigator.clipboard.writeText(text); toast('コピーしました', 'ok', 1500); } catch {
    const ta = el('textarea', {}, text);
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); toast('コピーしました', 'ok', 1500); } catch { toast('コピーできませんでした', 'error'); }
    ta.remove();
  }
}

function fmtMB(bytes) {
  return `${(Number(bytes || 0) / 1024 / 1024).toFixed(1)}MB`;
}

function detailMaterials(s) {
  const card = el('section', { class: 'card' });
  const input = el('input', {
    type: 'file', multiple: true, accept: '.pdf,.png,.jpg,.jpeg,.webp,.txt,.md,.csv',
    onchange: async (e) => {
      const files = Array.from(e.target.files || []);
      for (const f of files) {
        try { await pipe.addMaterial(s, f); } catch (err) { toast(errorMessage(err), 'error'); }
      }
      toast(`${files.length}件の資料を追加しました。②の実行時に読み込みます。`, 'ok');
      renderDetailBody();
    },
  });
  card.append(el('p', { class: 'muted small' }, 'PDF・画像・テキストに対応しています。議事録の作成時に読み込み、専門用語は単語帳に自動で追加します。'), input);
  const box = el('div', { class: 'list' });
  card.appendChild(box);
  (async () => {
    const local = await getFiles(s.id);
    const names = new Set(local.map((f) => f.name));
    local.forEach((f) => box.appendChild(el('div', { class: 'list-item static' }, `📄 ${f.name}`, el('span', { class: 'muted small' }, `${fmtMB(f.size)}・${f.uploaded ? 'OneDrive保存済み' : '未送信'}`))));
    const totalBytes = local.reduce((a, f) => a + (f.size || 0), 0);
    if (totalBytes > 20 * 1024 * 1024) {
      card.insertBefore(el('p', { class: 'warn small' }, `資料の合計が${fmtMB(totalBytes)}です。PDF・画像は合計40MBまでAIに渡し、超えた分は省きます。ページ数の多いPDFは議事録1回の費用が増えるため、必要なページだけにすると確実です。`), box);
    }
    if (s.folderPath && S.status.ms) {
      const remote = await od.listChildren(od.joinPath(s.folderPath, '資料')).catch(() => []);
      remote.filter((r) => !names.has(r.name)).forEach((r) => box.appendChild(el('div', { class: 'list-item static' }, `📄 ${r.name}`, el('span', { class: 'muted small' }, 'OneDrive'))));
    }
    if (!box.childNodes.length) box.appendChild(el('p', { class: 'muted small' }, '資料はありません。'));
  })();
  return card;
}

async function renameSpeakersUI(s) {
  const names = [...new Set(s.transcript.segments.map((x) => x.speaker))];
  const r = await formDialog('話者名を変更', names.map((n) => ({ name: n, label: `${n} →`, value: n })));
  if (!r) return;
  try {
    await pipe.renameSpeakers(s, r);
    toast('話者名を変更しました（案件の話者として記憶しました）', 'ok');
    renderDetailBody();
  } catch (e) { toast(errorMessage(e), 'error'); }
}

async function calendarUI(s) {
  const st = getSettings();
  if (!st.google.clientId) {
    const ev = pipe.calendarEvent(s, s.minutes && s.minutes.data);
    window.open(tp.calendarTemplateUrl(ev, st.google.account), '_blank', 'noopener');
    return;
  }
  try {
    if (!(await googleTokenValid())) await googleLogin(st.google.clientId, st.google.account);
    S.status.google = true;
    if (!s.folderPath) await pipe.ensureFolder(s);
    if (!s.folderWebUrl && s.folderPath) s.folderWebUrl = await od.webUrl(s.folderPath);
    const ok = await pipe.tryCalendar(s, s.minutes && s.minutes.data);
    toast(ok ? 'カレンダーに登録しました' : 'カレンダーに登録できませんでした', ok ? 'ok' : 'warn');
  } catch (e) { toast(errorMessage(e), 'error'); }
}

// ===========================================================================
// 単語帳・改善記録（REQ-7 / REQ-11）
// ===========================================================================
function viewVocab() {
  const wrap = el('div', { class: 'col' });
  const projects = vocab.projects();
  if (!projects.find((p) => p.id === S.vocabProject)) S.vocabProject = 'common';
  const p = vocab.projectById(S.vocabProject);
  const sel = el('select', { class: 'input', onchange: (e) => { S.vocabProject = e.target.value; render(); } },
    projects.map((x) => el('option', { value: x.id, selected: x.id === p.id }, `${x.name}（${x.terms.filter((t) => !t.deleted).length}語）`)));
  const wrong = el('input', { class: 'input', placeholder: '聞き間違い・読み（例：ふぉるてっく）' });
  const correct = el('input', { class: 'input', placeholder: '正しい表記（例：フォルテック）' });
  const cat = el('input', { class: 'input', placeholder: '分類（人名・組織名など）' });
  const addTerm = async () => {
    if (!correct.value.trim()) { toast('正しい表記を入力してください', 'warn'); return; }
    const added = await vocab.addTerms(p.id, [{ wrong: wrong.value, correct: correct.value, category: cat.value }], '手入力');
    toast(added.length ? '追加しました' : '既に登録されています', added.length ? 'ok' : 'info');
    render();
  };
  const terms = p.terms.filter((t) => !t.deleted).slice().reverse();
  wrap.appendChild(section('案件別の単語帳',
    el('div', { class: 'row nowrap' }, sel,
      el('button', { class: 'btn', type: 'button', onclick: async () => { const r = await formDialog('案件を追加', [{ name: 'name', label: '案件名', placeholder: '例：【案件C】基幹システム刷新' }]); if (r && r.name.trim()) { const np = await vocab.addProject(r.name); S.vocabProject = np.id; render(); } } }, '＋案件'),
      p.id !== 'common' ? el('button', { class: 'btn ghost', type: 'button', onclick: async () => { const r = await formDialog('案件名を変更', [{ name: 'name', label: '案件名', value: p.name }]); if (r && r.name.trim()) { await vocab.renameProject(p.id, r.name); render(); } } }, '名前変更') : null,
      p.id !== 'common' ? el('button', { class: 'btn ghost', type: 'button', onclick: async () => { if (await confirmDialog(`案件「${p.name}」を削除しますか？`, { ok: '削除', danger: true })) { await vocab.deleteProject(p.id); S.vocabProject = 'common'; render(); } } }, '削除') : null),
    el('p', { class: 'muted small' }, '「全案件共通」はすべての会議に、各案件の単語帳はその案件を選んだ会議に使います。2台の端末で自動的に共有されます。'),
    el('div', { class: 'row' }, wrong, correct, cat, el('button', { class: 'btn primary', type: 'button', onclick: addTerm }, '追加')),
    el('div', { class: 'table-wrap' }, el('table', {},
      el('thead', {}, el('tr', {}, el('th', {}, '正しい表記'), el('th', {}, '聞き間違い・読み'), el('th', {}, '分類'), el('th', {}, '追加元'), el('th', {}, ''))),
      el('tbody', {}, terms.length ? terms.map((t) => el('tr', {},
        el('td', {}, t.correct), el('td', {}, t.wrong || ''), el('td', {}, t.category || ''),
        el('td', {}, t.source === '手入力' ? '手入力' : chip(`自動：${t.source}`, 'muted')),
        el('td', {}, el('button', { class: 'icon-btn', type: 'button', title: '削除', onclick: async () => { await vocab.removeTerm(p.id, t.id); render(); } }, '×'))))
        : el('tr', {}, el('td', { colspan: '5', class: 'muted' }, 'まだ登録はありません。')))))));

  const speakers = vocab.speakersFor(p.id);
  wrap.appendChild(section('話者（名前・役割）',
    el('p', { class: 'muted small' }, '書き起こしで話者を推定するときのヒントに使います。「話者名を変更」で付けた名前は自動で登録されます。'),
    el('div', { class: 'list' }, speakers.length ? speakers.map((sp) => el('div', { class: 'list-item static' },
      el('span', {}, `${sp.name}${sp.role ? `（${sp.role}）` : ''}`),
      el('button', { class: 'btn small ghost', type: 'button', onclick: async () => { const r = await formDialog('役割を設定', [{ name: 'role', label: `${sp.name}の役割`, value: sp.role || '', placeholder: '例：進行役、先方担当者' }]); if (r) { await vocab.updateSpeaker(p.id, sp.id, { role: r.role.trim() }); render(); } } }, '役割'),
      el('button', { class: 'icon-btn', type: 'button', onclick: async () => { await vocab.updateSpeaker(p.id, sp.id, { deleted: true }); render(); } }, '×')))
      : el('p', { class: 'muted small' }, 'まだ登録はありません。')),
    el('button', { class: 'btn small', type: 'button', onclick: async () => { const r = await formDialog('話者を追加', [{ name: 'name', label: '名前' }, { name: 'role', label: '役割（任意）' }]); if (r && r.name.trim()) { await vocab.rememberSpeakers(p.id, [r.name], r.role.trim()); render(); } } }, '＋話者')));

  const rules = p.rules.filter((r) => !r.deleted);
  wrap.appendChild(section('議事録の書き方のルール',
    el('p', { class: 'muted small' }, '議事録の手直しから自動で学びます。手で追加することもできます。'),
    el('div', { class: 'list' }, rules.length ? rules.map((r) => el('div', { class: 'list-item static' }, el('span', {}, r.text),
      r.source !== '手入力' ? chip(`自動：${r.source}`, 'muted') : null,
      el('button', { class: 'icon-btn', type: 'button', onclick: async () => { await vocab.removeRule(p.id, r.id); render(); } }, '×')))
      : el('p', { class: 'muted small' }, 'まだルールはありません。')),
    el('button', { class: 'btn small', type: 'button', onclick: async () => { const r = await formDialog('ルールを追加', [{ name: 'text', label: 'ルール', placeholder: '例：先方の発言は「先方」と表記する' }]); if (r && r.text.trim()) { await vocab.addRules(p.id, [r.text]); render(); } } }, '＋ルール')));

  const usageBox = el('div');
  const logBox = el('div', { class: 'list' });
  wrap.appendChild(section('利用状況と改善の記録', usageBox, el('h3', {}, '改善の記録'), logBox));
  (async () => {
    const usage = await getUsage();
    const m = tp.monthlySummary(usage);
    const budget = Number(getSettings().budget.monthlyUsd) || 10;
    const ratio = Math.min(1, m.cost / budget);
    usageBox.append(
      el('p', {}, `今月のAI利用：${m.calls}回、音声${Math.round(m.audioSec / 60)}分、推定費用 約${m.cost.toFixed(2)}ドル（予算${budget}ドルの${Math.round(ratio * 100)}%）`),
      progressBar(ratio),
      el('p', { class: 'muted small' }, '推定費用は公開単価からの概算です。正確な金額はGoogle Cloudの請求画面で確認してください。'));
    const log = await getKaizen();
    if (!log.length) logBox.appendChild(el('p', { class: 'muted small' }, 'まだ記録はありません。'));
    log.slice(0, 50).forEach((x) => logBox.appendChild(el('div', { class: 'list-item static' }, el('span', { class: 'muted small mono' }, fmtDateTime(x.at)), el('span', {}, x.text))));
  })();
  return wrap;
}

// ===========================================================================
// 設定（REQ-8 / REQ-12）
// ===========================================================================
function viewSettings() {
  const st = getSettings();
  const ss = setupStatus(st);
  const wrap = el('div', { class: 'col' });
  const f = (label, input, help) => el('label', { class: 'field' }, el('span', {}, label), input, help ? el('small', { class: 'muted' }, help) : null);
  const inp = (value, attrs = {}) => el('input', { class: 'input', value: value ?? '', ...attrs });

  wrap.appendChild(section('初回設定の状況',
    el('ol', { class: 'steps' },
      el('li', { class: ss.onedrive && S.status.ms ? 'done' : '' }, 'OneDrive：アプリIDを入力してサインイン'),
      el('li', { class: ss.gemini ? 'done' : '' }, 'Gemini：課金を有効にしたプロジェクトのAPIキーを入力し、確認欄にチェック'),
      el('li', { class: ss.google ? 'done' : '' }, 'Googleカレンダー：クライアントIDを入力（任意。未設定でも登録リンクで代替）')),
    redirectBox(),
    el('p', { class: 'muted small' }, '設定は端末ごとに保存されます。2台目の端末では、OneDriveにサインインすると単語帳と共通の設定が自動で取り込まれます（APIキーは各端末で入力）。')));

  // OneDrive
  const odClient = inp(st.onedrive.clientId, { placeholder: '00000000-0000-0000-0000-000000000000' });
  const odRoot = inp(st.onedrive.rootPath);
  const odStatus = el('div', { class: 'od-status' });
  wrap.appendChild(section('OneDrive',
    f('アプリID（クライアントID）', odClient, 'Microsoft Entra でこのアプリを登録したときの「アプリケーション（クライアント）ID」です（手順書2章）。'),
    f('保存先フォルダ', odRoot, 'OneDriveの一番上（「自分のファイル」）からのパス。サインインに成功した時点で作られ、中に「年月日_時分」のフォルダを作ります。'),
    el('div', { class: 'row' },
      el('button', { class: 'btn primary', type: 'button', onclick: async () => {
        await saveSettings({ onedrive: { clientId: odClient.value.trim(), rootPath: normalizeRoot(odRoot.value) } });
        odRoot.value = rootPath();
        await signInOneDrive();
      } }, S.status.ms ? '保存して再サインイン' : '保存してサインイン'),
      el('button', { class: 'btn', type: 'button', onclick: async () => { await saveSettings({ onedrive: { clientId: odClient.value.trim(), rootPath: normalizeRoot(odRoot.value) } }); odRoot.value = rootPath(); toast('保存しました（OneDriveのフォルダは、「保存してサインイン」に成功した時点で作られます）', 'ok', 6000); renderHeader(); } }, '保存のみ'),
      S.status.ms ? el('button', { class: 'btn ghost', type: 'button', onclick: async () => { await msLogout(); S.status.ms = false; watcher.stop(); render(); } }, 'サインアウト') : null),
    odStatus));
  renderOdStatus(odStatus);

  // Gemini
  const key = inp(st.gemini.apiKey, { type: 'password', autocomplete: 'off', placeholder: 'AQ.…（AI Studioで作ったキー）' });
  const paid = el('input', { type: 'checkbox', checked: st.gemini.paidConfirmed });
  const tModel = inp(st.gemini.transcribeModel);
  const xModel = inp(st.gemini.textModel);
  const mModel = inp(st.gemini.minutesModel);
  wrap.appendChild(section('Gemini（書き起こし・議事録）',
    f('APIキー', key, 'この端末にだけ保存されます。AI Studioで作ったキーは、最初からGemini APIだけに制限されています。月の上限（Spend）を設定し、自動チャージはオフにしてください（手順書3章）。'),
    el('label', { class: 'check strong' }, paid, 'このキーは、課金（Cloud Billing）を有効にしたプロジェクトのものです'),
    el('p', { class: 'muted small' }, '無料枠では、送った内容がGoogleの製品改善に使われ、人が読むこともあります。課金を有効にしたプロジェクトでは学習に使われません。確認するまで、このアプリはGeminiに何も送りません。'),
    el('p', { class: 'muted small' }, '確認のしかた：Google AI Studio の「Billing」画面で、このキーのプロジェクトが有料（Paid）になっていて、前払い残高が0ドルより多いこと。前払い残高は購入から12か月で失効し、0ドルになるとGoogle AI Proのクレジットも使えなくなります（手順書3章）。'),
    el('details', {}, el('summary', {}, 'モデルの設定（通常は変更不要）'),
      f('書き起こし（話者分離）', tModel), f('用語補正・単語の抽出', xModel), f('議事録', mModel)),
    el('div', { class: 'row' },
      el('button', { class: 'btn primary', type: 'button', onclick: async () => { await saveSettings({ gemini: { apiKey: key.value.trim(), paidConfirmed: paid.checked, transcribeModel: tModel.value.trim(), textModel: xModel.value.trim(), minutesModel: mModel.value.trim() } }); toast('保存しました', 'ok'); renderHeader(); } }, '保存'),
      el('button', { class: 'btn', type: 'button', onclick: async () => { try { const models = await testApiKey(key.value.trim()); const need = [tModel.value.trim(), xModel.value.trim(), mModel.value.trim()]; const missing = need.filter((m) => !models.includes(m)); toast(missing.length ? `キーは有効ですが、次のモデルが一覧にありません：${missing.join('、')}` : 'キーは有効です。必要なモデルも利用できます。', missing.length ? 'warn' : 'ok', 8000); } catch (e) { toast(errorMessage(e), 'error'); } } }, 'キーを確認'))));

  // Google
  const gClient = inp(st.google.clientId, { placeholder: '….apps.googleusercontent.com' });
  const gAcc = inp(st.google.account, { placeholder: '例：自分のGmailアドレス', type: 'email' });
  const gAuto = el('input', { type: 'checkbox', checked: st.google.autoCalendar });
  const gSum = el('input', { type: 'checkbox', checked: st.privacy.calendarSummary });
  wrap.appendChild(section('Googleカレンダー',
    f('OAuthクライアントID', gClient, '未設定の場合は、会議ごとに「カレンダー登録」で登録画面を開く方式になります。'),
    f('登録するアカウント', gAcc),
    el('label', { class: 'check' }, gAuto, '録音した会議を自動で登録する'),
    el('label', { class: 'check' }, gSum, 'カレンダーに要約・決定事項・ToDoを書く（「機密」の会議は書きません）'),
    el('div', { class: 'row' },
      el('button', { class: 'btn primary', type: 'button', onclick: async () => { await saveSettings({ google: { clientId: gClient.value.trim(), account: gAcc.value.trim(), autoCalendar: gAuto.checked }, privacy: { calendarSummary: gSum.checked } }); await pushShared(); if (gClient.value.trim()) { try { await googleLogin(gClient.value.trim(), gAcc.value.trim()); S.status.google = true; toast('Googleカレンダーを連携しました', 'ok'); } catch (e) { toast(errorMessage(e), 'error'); } } render(); } }, '保存して連携'),
      st.google.clientId ? el('button', { class: 'btn ghost', type: 'button', onclick: async () => { await googleLogout(); S.status.google = false; render(); } }, '連携を解除') : null)));

  // ほかの端末（スマホなど）を設定する
  const withKey = el('input', { type: 'checkbox', checked: true });
  wrap.appendChild(section('ほかの端末（スマホなど）を設定する',
    el('p', { class: 'small' }, 'この端末の設定（OneDrive・Gemini・Googleカレンダー）を、QRコードでほかの端末に渡します。ほかの端末では、カメラで読み取って「取り込む」を押し、OneDriveにサインインするだけで使えます。'),
    el('label', { class: 'check' }, withKey, 'GeminiのAPIキーも渡す（外すと、ほかの端末でキーを貼り付けます）'),
    el('div', { class: 'row' },
      el('button', { class: 'btn primary', type: 'button', onclick: () => showSetupQr(withKey.checked) }, 'QRコードを表示'))));

  // 録音・書き起こし・メモ
  const r = st.recording;
  const cb = (v) => el('input', { type: 'checkbox', checked: v });
  const hp = cb(r.highpass); const ns = cb(r.noiseSuppression); const agc = cb(r.autoGainControl); const ec = cb(r.echoCancellation); const hold = cb(r.holdToStop); const auto1 = cb(r.autoStep1);
  const filler = el('select', { class: 'input' }, el('option', { value: 'standard', selected: st.transcript.fillerLevel === 'standard' }, '標準（えー・あのー等を除く）'), el('option', { value: 'none', selected: st.transcript.fillerLevel === 'none' }, 'そのまま（一字一句）'));
  const phrases = el('textarea', { class: 'input', rows: 4 }, (st.memo.phrases || []).join('\n'));
  const devName = inp(st.deviceName);
  const poll = inp(st.sync.pollSec, { type: 'number', min: '3', max: '60' });
  wrap.appendChild(section('録音・書き起こし・メモ',
    f('この端末の名前', devName, 'もう一方の端末に「○○で録音中」と表示されます。'),
    el('label', { class: 'check' }, hp, '低い雑音（空調・机の振動）を除く'),
    el('label', { class: 'check' }, ns, 'ブラウザの雑音抑制を使う'),
    el('label', { class: 'check' }, agc, '音量を自動で調整する'),
    el('label', { class: 'check' }, ec, 'エコー除去（スピーカーで会議音声を流しながら同じ端末で録る場合のみ）'),
    el('label', { class: 'check' }, hold, '録音の終了は長押しで行う（誤タップ防止）'),
    el('label', { class: 'check' }, auto1, '録音を終了したら、自動で①（保存・書き起こし）を始める'),
    f('ケバ取り', filler),
    f('定型句（1行に1つ）', phrases),
    f('2台の同期の間隔（秒）', poll),
    el('button', {
      class: 'btn primary', type: 'button',
      onclick: async () => {
        await saveSettings({
          deviceName: devName.value.trim() || st.deviceName,
          recording: { highpass: hp.checked, noiseSuppression: ns.checked, autoGainControl: agc.checked, echoCancellation: ec.checked, holdToStop: hold.checked, autoStep1: auto1.checked },
          transcript: { fillerLevel: filler.value },
          memo: { phrases: phrases.value.split('\n').map((x) => x.trim()).filter(Boolean) },
          sync: { pollSec: Math.max(3, Number(poll.value) || 5) },
        });
        await pushShared();
        toast('保存しました', 'ok');
      },
    }, '保存')));

  // 機密保護・予算
  const ret = inp(st.privacy.localRetentionDays, { type: 'number', min: '0', max: '90' });
  const budget = inp(st.budget.monthlyUsd, { type: 'number', min: '1', step: '1' });
  wrap.appendChild(section('機密保護・予算',
    f('端末内の録音データを残す日数', ret, 'OneDriveへの保存を確認した録音データは、この日数を過ぎると端末から削除します（0で自動削除しない）。'),
    f('月の予算（ドル）', budget, 'Google AI Proのクレジットは月10ドルです。単語帳タブで消費の目安を表示します。'),
    el('div', { class: 'row' },
      el('button', { class: 'btn primary', type: 'button', onclick: async () => { await saveSettings({ privacy: { localRetentionDays: Math.max(0, Number(ret.value) || 0) }, budget: { monthlyUsd: Math.max(1, Number(budget.value) || 10) } }); await pushShared(); toast('保存しました', 'ok'); } }, '保存'),
      el('button', { class: 'btn', type: 'button', onclick: exportSettings }, '設定を書き出す'),
      el('label', { class: 'btn' }, '設定を読み込む', el('input', { type: 'file', accept: 'application/json', class: 'hidden', onchange: importSettings })),
      el('button', { class: 'btn danger', type: 'button', onclick: wipeLocal }, '端末内のデータを消去'))));

  wrap.appendChild(el('p', { class: 'muted small center' }, `端末ID：${st.deviceId}　／　AI会議録音＆OneDrive自動議事録 v2.2`));
  return wrap;
}

async function pushShared() {
  if (!S.status.ms) return;
  try { await od.uploadJson(od.joinPath(rootPath(), '_app', 'settings.json'), sharedPart()); } catch { /* 次回 */ }
}

async function exportSettings() {
  const withKey = await confirmDialog('書き出すファイルにGeminiのAPIキーを含めますか？（含めた場合は、ファイルの扱いに注意してください）', { ok: '含める', cancel: '含めない' });
  const st = JSON.parse(JSON.stringify(getSettings()));
  delete st.deviceId; delete st.deviceName;
  if (!withKey) st.gemini.apiKey = '';
  const blob = new Blob([JSON.stringify({ app: 'mtgrec', version: 2, exportedAt: Date.now(), settings: st, vocab: vocab.getVocab() }, null, 1)], { type: 'application/json' });
  const a = el('a', { href: URL.createObjectURL(blob), download: `会議録音アプリ_設定_${new Date().toISOString().slice(0, 10)}.json` });
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

async function importSettings(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  try {
    const j = JSON.parse(await file.text());
    if (j.app !== 'mtgrec') throw new Error('このアプリの設定ファイルではありません');
    const s = j.settings || {};
    delete s.deviceId; delete s.deviceName;
    if (s.gemini && !s.gemini.apiKey) delete s.gemini.apiKey;
    await saveSettings(s);
    if (j.vocab) await vocab.mergeRemote(j.vocab);
    toast('設定を読み込みました', 'ok');
    render();
  } catch (err) { toast(`読み込めませんでした：${errorMessage(err)}`, 'error'); }
}

async function wipeLocal() {
  if (S.rec) { toast('録音中は消去できません', 'warn'); return; }
  if (!(await confirmDialog('この端末内の録音データ・会議・単語帳・設定・サインイン情報をすべて消去します。OneDriveに保存済みのデータは消えません。よろしいですか？', { ok: '消去する', danger: true }))) return;
  await wipeAll();
  try { localStorage.clear(); sessionStorage.clear(); } catch { /* 無視 */ }
  if ('caches' in window) { for (const k of await caches.keys()) await caches.delete(k); }
  toast('消去しました。再読み込みします。', 'ok');
  await sleep(800);
  location.reload();
}

boot().catch((e) => {
  console.error(e);
  const main = document.getElementById('main');
  if (main) main.textContent = `起動できませんでした：${errorMessage(e)}`;
});
