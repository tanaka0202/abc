// 処理の流れ：①録音保存＆書き起こし作成、②メモを組み込んで議事録作成＆保存（REQ-5）
// 各手順の完了を session.steps に記録し、失敗しても「再試行」で続きから再開できる。
import * as od from './onedrive.js';
import * as ai from './gemini.js';
import { upsertEvent } from './gcal.js';
import { googleTokenValid } from './auth.js';
import { getChunks, putSession, getSession, listSessions, getFiles, putFile, deleteChunks } from './db.js';
import { getSettings } from './settings.js';
import { rootPath, metaPath, readAllMemos } from './sync.js';
import * as tp from './textproc.js';
import * as vocab from './vocab.js';
import { getMyMemos, flushMemos, combine } from './memos.js';
import { pcmBlobFromChunks, wavFromPcmBlob, wavSlice, parseWav, SAMPLE_RATE } from './wav.js';
import { logKaizen } from './usage.js';
import { AppError, folderNameFor, fmtDateTime, fmtTime, blobToBase64, safeFileName, uuid, errorMessage } from './util.js';

const indexPath = () => od.joinPath(rootPath(), '_app', 'index.json');

/** OneDriveの会議フォルダを用意する */
export async function ensureFolder(session) {
  if (session.folderPath) return session;
  const created = await od.createUniqueFolder(rootPath(), folderNameFor(session.startedAt || session.createdAt || Date.now()));
  session.folderPath = created.path;
  session.folderName = created.name;
  try { session.folderWebUrl = await od.webUrl(created.path); } catch { /* 後で取得 */ }
  session.steps = { ...(session.steps || {}), folder: true };
  await putSession(session);
  return session;
}

/** 他の端末が見る会議の状態（_meta/session.json） */
export function sessionMeta(s) {
  return {
    id: s.id, folderName: s.folderName, folderPath: s.folderPath, folderWebUrl: s.folderWebUrl,
    title: s.title, projectId: s.projectId, templateId: s.templateId, participants: s.participants || [],
    confidential: !!s.confidential, startedAt: s.startedAt, endedAt: s.endedAt, durationSec: s.durationSec,
    status: s.status, steps: s.steps || {}, progress: s.progress || null, error: s.error || null,
    eventId: s.eventId || null, deviceId: s.deviceId, deviceName: s.deviceName, gaps: s.gaps || [],
    materials: (s.materials || []).map((m) => ({ id: m.id, name: m.name, type: m.type, size: m.size })),
    updatedAt: Date.now(),
  };
}

export async function writeMeta(session) {
  if (!session.folderPath) return;
  try { await od.uploadJson(metaPath(session.folderPath, 'session.json'), sessionMeta(session)); } catch (e) { console.warn(e); }
}

let lastMetaWrite = 0;
async function progress(session, onProgress, step, text, ratio = null) {
  session.progress = { step, text, ratio, at: Date.now() };
  if (onProgress) onProgress(session.progress);
  if (Date.now() - lastMetaWrite > 8000 || ratio === 1) { lastMetaWrite = Date.now(); writeMeta(session); }
}

/** 録音データ（PCM）を取得：端末内にあればそれを、なければOneDriveのWAVを使う */
async function getPcm(session) {
  const chunks = await getChunks(session.id);
  if (chunks.length) return pcmBlobFromChunks(chunks);
  if (!session.folderPath) throw new AppError('録音データが見つかりません。');
  const res = await od.download(od.joinPath(session.folderPath, '録音データ.wav'), 'blob');
  if (!res) throw new AppError('OneDriveに録音データが見つかりません。');
  const { pcm, sampleRate } = await parseWav(res.body);
  if (sampleRate !== SAMPLE_RATE) throw new AppError('このアプリで録音したWAV（16kHz）以外は再処理できません。');
  return pcm;
}

// ===========================================================================
// ① 録音保存 ＆ 書き起こし作成
// ===========================================================================
export async function runStep1(session, { onProgress, signal } = {}) {
  session.status = 'processing';
  session.error = null;
  session.steps = session.steps || {};
  await putSession(session);
  try {
    // 1. 保存先フォルダ
    await progress(session, onProgress, 'folder', 'OneDriveに保存先フォルダを用意しています');
    await ensureFolder(session);
    await writeMeta(session);

    const pcm = await getPcm(session);
    if (!session.durationSec) session.durationSec = pcm.size / 2 / SAMPLE_RATE;
    if (!session.endedAt && session.startedAt) session.endedAt = session.startedAt + session.durationSec * 1000;

    // 2. 録音データ
    if (!session.steps.audio) {
      await progress(session, onProgress, 'audio', '録音データを保存しています', 0);
      await od.upload(od.joinPath(session.folderPath, '録音データ.wav'), wavFromPcmBlob(pcm), 'audio/wav',
        (r) => progress(session, onProgress, 'audio', `録音データを保存しています（${Math.round(r * 100)}%）`, r));
      await od.uploadJson(metaPath(session.folderPath, 'levels.json'), { levels: session.levels || [], gaps: session.gaps || [] });
      session.steps.audio = true;
      session.audioSavedAt = Date.now();
      await putSession(session);
    }
    await flushMemos(session.id, session.folderPath).catch(() => {});
    await uploadPendingMaterials(session);

    // 3. 書き起こし
    if (!session.steps.transcript) {
      await transcribe(session, pcm, { onProgress, signal });
      session.steps.transcript = true;
      await putSession(session);
    }

    // 4. カレンダー（サインインが有効な場合のみ。切れていれば②で登録）
    await progress(session, onProgress, 'calendar', 'カレンダーに登録しています');
    await tryCalendar(session).catch((e) => { session.calendarError = errorMessage(e); });

    session.status = session.steps.minutes ? 'done' : 'transcribed';
    session.progress = null;
    await putSession(session);
    await writeMeta(session);
    await updateIndex(session);
    await logKaizen('session', `①完了：「${session.title}」（${fmtTime(session.durationSec)}、発言${(session.transcript?.segments || []).length}件）`);
    return session;
  } catch (e) {
    session.status = 'error';
    session.error = errorMessage(e);
    await putSession(session);
    await writeMeta(session);
    throw e;
  } finally {
    if (onProgress) onProgress(null);
  }
}

const OVERLAP_SEC = 15;

/** 書き起こし専用モデルが使えない（未提供・仕様変更）と判断できるエラーか */
const modelUnavailable = (e) => e && e.status >= 400 && e.status < 500 && ![401, 402, 403, 429].includes(e.status);
/** 再試行・代替で回復しうる一時的な失敗か（中身が空・理由不明の拒否） */
const transient = (e) => e && (e.code === 'empty' || e.code === 'blocked');

/**
 * 1区切りを書き起こす。
 * Googleの書き起こしモデルは、まれに「成功」の応答で中身が空になる不具合が報告されているため、
 * 空なら再試行し、それでも空なら汎用モデルで代替する。ほぼ無音の区切りは空でも正常とみなす。
 */
async function transcribeSlice(wav, opts, { silent, state, session }) {
  if (!state.useGeneral) {
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const segs = await ai.transcribeChunk(wav, opts);
        if (segs.length || silent) return segs;
      } catch (e) {
        if (modelUnavailable(e)) {
          state.useGeneral = true;
          await logKaizen('fallback', `書き起こし専用モデルが使えなかったため汎用モデルで代替（${e.apiMessage || e.message}）`);
          break;
        }
        if (!transient(e)) throw e;
      }
      state.emptyRetries++;
    }
    if (!state.useGeneral) await logKaizen('fallback', `「${session.title}」の書き起こし（区切り${opts.chunkIndex + 1}）が空で返ったため、汎用モデルで代替しました`);
  }
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const segs = await ai.transcribeChunkGeneral(wav, opts);
      if (segs.length || silent) return segs;
    } catch (e) {
      if (!transient(e)) throw e;
    }
    state.emptyRetries++;
  }
  if (silent) return [];
  throw new AppError('書き起こしの結果が空で返りました（GoogleのAPIの一時的な不具合の可能性があります）。しばらくしてから「① を再試行」を押してください。録音データは保存済みです。');
}

async function transcribe(session, pcm, { onProgress, signal }) {
  const st = getSettings();
  const duration = pcm.size / 2 / SAMPLE_RATE;
  // 区切りは最長10分（スマホのメモリ対策）。前後15秒を重ねて切り出し、話者をそろえる
  const chunkSec = Math.max(1, Math.min(25, Number(st.gemini.chunkMinutes) || 10)) * 60;
  const plan = tp.planSlices(duration, chunkSec, session.levels || [], OVERLAP_SEC);
  const terms = vocab.termsFor(session.projectId);
  const diarization = session.diarization !== false && st.transcript.diarization !== false;
  const state = { useGeneral: false, emptyRetries: 0 };
  const results = [];
  for (let i = 0; i < plan.length; i++) {
    const c = plan[i];
    await progress(session, onProgress, 'transcribe', `書き起こしています（${i + 1}/${plan.length}）`, i / plan.length);
    const wav = wavSlice(pcm, c.sliceStart, c.sliceEnd);
    const opts = {
      chunkIndex: i, offsetSec: c.sliceStart, durSec: c.sliceEnd - c.sliceStart, diarization,
      vocab: terms.map((t) => t.correct), participants: session.participants, sessionId: session.id, signal,
    };
    const silent = tp.isSilent(session.levels, c.sliceStart, c.sliceEnd);
    const segs = await transcribeSlice(wav, opts, { silent, state, session });
    results.push({ ...c, segs });
  }
  const stitched = tp.stitchChunks(results);
  let segs = tp.mergeSegments(stitched.segments.map((x, i) => ({ ...x, id: i })));
  session.speakerLinks = stitched.linked;
  const useGeneral = state.useGeneral;

  // 用語の補正と話者のそろえ（失敗しても書き起こし自体は残す）
  await progress(session, onProgress, 'refine', '単語帳を使って用語を補正しています', 0.9);
  let mapping = {};
  let applied = 0;
  try {
    const project = vocab.projectById(session.projectId);
    const notes = vocab.speakersFor(session.projectId).map((s) => `${s.name}${s.role ? `（${s.role}）` : ''}`).join('、');
    const res = await ai.refineTranscript(segs, {
      terms, participants: session.participants, speakerNotes: notes, sessionId: session.id, signal,
    });
    mapping = res.mapping || {};
    const r = tp.applyCorrections(segs, res.corrections);
    segs = r.segments;
    applied = r.applied;
    if (project) session.projectName = project.name;
  } catch (e) {
    session.refineError = errorMessage(e);
  }
  segs = tp.assignSpeakers(segs, mapping);
  if (st.transcript.fillerLevel === 'standard') segs = segs.map((s) => ({ ...s, text: tp.removeFillers(s.text) }));

  session.transcript = {
    segments: segs, mapping, corrections: applied, createdAt: Date.now(),
    model: useGeneral ? st.gemini.textModel : st.gemini.transcribeModel,
    chunks: plan.length, speakerLinks: stitched.linked, emptyRetries: state.emptyRetries,
  };
  await putSession(session);
  await saveTranscriptFiles(session);
}

export async function saveTranscriptFiles(session) {
  const segs = session.transcript.segments;
  const meta = { title: session.title, startedAt: session.startedAt };
  await od.uploadText(od.joinPath(session.folderPath, '書き起こし.txt'), tp.segmentsToTxt(segs, meta));
  await od.uploadText(od.joinPath(session.folderPath, '字幕.srt'), tp.segmentsToSrt(segs));
  await od.uploadJson(metaPath(session.folderPath, 'transcript.json'), session.transcript);
}

// ===========================================================================
// カレンダー（REQ-9）
// ===========================================================================
export function calendarEvent(session, minutesData) {
  const st = getSettings();
  const lines = [`会議名：${session.title}`, `日時：${fmtDateTime(session.startedAt)}（約${Math.max(1, Math.round((session.durationSec || 0) / 60))}分）`];
  if (session.folderWebUrl) lines.push(`保存先（OneDrive）：${session.folderWebUrl}`);
  const allowSummary = st.privacy.calendarSummary && !session.confidential;
  if (minutesData && allowSummary) {
    lines.push('', `■要約\n${minutesData.summary || ''}`);
    if (minutesData.decisions && minutesData.decisions.length) lines.push('', '■決定事項', ...minutesData.decisions.slice(0, 10).map((d) => `・${d.text}`));
    if (minutesData.todos && minutesData.todos.length) lines.push('', '■ToDo', ...minutesData.todos.slice(0, 10).map((t) => `・${t.owner || '未定'}：${t.task}（${t.due || '期限未定'}）`));
  } else if (session.confidential) {
    lines.push('', '（機密の会議のため、要約は記載していません）');
  }
  return {
    uid: `${session.id}@mtgrec`,
    title: `【会議記録】${session.title}`,
    start: session.startedAt,
    end: session.endedAt || (session.startedAt + (session.durationSec || 0) * 1000),
    url: session.folderWebUrl || '',
    description: lines.join('\n').slice(0, 7000),
  };
}

/** 予定を登録・更新する。サインインが切れている場合は保留にして true/false を返す */
export async function tryCalendar(session, minutesData = null) {
  const st = getSettings();
  const ev = calendarEvent(session, minutesData);
  session.calendarUrl = tp.calendarTemplateUrl(ev, st.google.account);
  if (!st.google.clientId || !st.google.autoCalendar) { await putSession(session); return false; }
  if (!(await googleTokenValid())) { session.calendarPending = true; await putSession(session); return false; }
  const res = await upsertEvent(session.id, ev);
  session.eventId = res && res.id;
  session.calendarPending = false;
  session.steps = { ...(session.steps || {}), calendar: true };
  await putSession(session);
  return true;
}

// ===========================================================================
// ② メモを組み込んで議事録作成 ＆ 保存
// ===========================================================================
export async function runStep2(session, { onProgress, signal } = {}) {
  session.error = null;
  try {
    await progress(session, onProgress, 'minutes', '書き起こしとメモを集めています', 0.05);
    if (!session.transcript) session.transcript = await od.readJson(metaPath(session.folderPath, 'transcript.json'));
    if (!session.transcript || !session.transcript.segments) throw new AppError('書き起こしがまだありません。先に①を実行してください。');
    const segs = session.transcript.segments;

    const memos = await collectMemos(session);
    const memoLines = memos.map((m) => `[${Number.isFinite(m.offset) ? (m.offset < 0 ? '開始前' : fmtTime(m.offset)) : '時刻なし'}] ${m.kind === 'star' ? '★' : ''}${m.text}`);

    await progress(session, onProgress, 'minutes', '会議資料を読み込んでいます', 0.15);
    const materials = await collectMaterials(session);
    const carry = await previousTodos(session);
    const project = vocab.projectById(session.projectId);
    const tpl = tp.templateById(session.templateId);

    await progress(session, onProgress, 'minutes', '議事録を作成しています（1〜2分かかります）', 0.3);
    const minutesArgs = {
      sessionId: session.id, title: session.title, templateId: session.templateId,
      when: `${fmtDateTime(session.startedAt)}〜（約${Math.max(1, Math.round((session.durationSec || 0) / 60))}分）`,
      participants: session.participants, projectName: project && project.id !== vocab.COMMON_ID ? project.name : '',
      terms: vocab.termsFor(session.projectId), rules: vocab.rulesFor(session.projectId),
      carryover: carry, memos: memoLines,
      transcriptText: segs.map((s) => `[${fmtTime(s.start)}] ${s.speaker}: ${s.text}`).join('\n'),
      materials, signal,
    };
    let raw;
    try {
      raw = await ai.generateMinutes(minutesArgs);
    } catch (e) {
      if (e.code !== 'empty' && e.code !== 'blocked') throw e;
      await progress(session, onProgress, 'minutes', '結果が空だったため、もう一度作成しています', 0.5);
      raw = await ai.generateMinutes(minutesArgs);
    }
    const { minutes, flagged, weak } = tp.verifyEvidence(raw, segs);
    const meta = {
      title: session.title, startedAt: session.startedAt, endedAt: session.endedAt, durationSec: session.durationSec,
      participants: session.participants && session.participants.length ? session.participants : null,
      projectName: project && project.id !== vocab.COMMON_ID ? project.name : '', templateName: tpl.name,
      folderUrl: session.folderWebUrl, materials: materials.map((m) => m.name),
    };
    const markdown = tp.minutesToMarkdown(minutes, meta);
    session.minutes = { data: minutes, markdown, flagged, weak, createdAt: Date.now(), edited: false };
    session.memosMerged = memos;
    await putSession(session);

    await progress(session, onProgress, 'save', '議事録・メモを保存しています', 0.85);
    await saveMinutesFiles(session, memos);

    await progress(session, onProgress, 'calendar', 'カレンダーを更新しています', 0.92);
    await tryCalendar(session, minutes).catch((e) => { session.calendarError = errorMessage(e); });

    // 資料から用語を学ぶ（REQ-7）
    const newMaterials = materials.filter((m) => !(session.learnedMaterials || []).includes(m.name));
    if (newMaterials.length) {
      try {
        const existing = vocab.termsFor(session.projectId).map((t) => t.correct);
        const res = await ai.extractTermsFromMaterials(newMaterials, existing, { sessionId: session.id });
        const added = await vocab.addTerms(session.projectId || vocab.COMMON_ID, res.terms || [], '資料');
        session.learnedMaterials = [...(session.learnedMaterials || []), ...newMaterials.map((m) => m.name)];
        session.learnedTerms = (session.learnedTerms || 0) + added.length;
      } catch (e) { session.learnError = errorMessage(e); }
    }

    session.steps = { ...(session.steps || {}), minutes: true };
    session.status = 'done';
    session.progress = null;
    await putSession(session);
    await writeMeta(session);
    await updateIndex(session);
    await logKaizen('session', `②完了：「${session.title}」（決定${minutes.decisions.length}件、ToDo${minutes.todos.length}件${flagged ? `、根拠未確認${flagged}件` : ''}）`);
    return session;
  } catch (e) {
    session.error = errorMessage(e);
    await putSession(session);
    await writeMeta(session);
    throw e;
  } finally {
    if (onProgress) onProgress(null);
  }
}

export async function saveMinutesFiles(session, memos) {
  const fp = session.folderPath;
  const meta = { title: session.title, startedAt: session.startedAt };
  await od.uploadText(od.joinPath(fp, '議事録.md'), session.minutes.markdown, 'text/markdown;charset=utf-8');
  if (memos) await od.uploadText(od.joinPath(fp, 'メモ.txt'), tp.memosToTxt(memos, meta));
  await od.uploadText(od.joinPath(fp, 'ToDo一覧.csv'), tp.todosToCsv(session.minutes.data.todos, meta), 'text/csv;charset=utf-8');
  await od.uploadText(od.joinPath(fp, 'カレンダー登録.ics'), tp.buildIcs(calendarEvent(session, session.minutes.data)), 'text/calendar;charset=utf-8');
  await od.uploadJson(metaPath(fp, 'minutes.json'), session.minutes);
}

/** 📌 メモだけ保存 */
export async function saveMemosOnly(session) {
  await ensureFolder(session);
  await flushMemos(session.id, session.folderPath);
  const memos = await collectMemos(session);
  await od.uploadText(od.joinPath(session.folderPath, 'メモ.txt'), tp.memosToTxt(memos, { title: session.title }));
  return memos;
}

export async function collectMemos(session) {
  const mine = await getMyMemos(session.id);
  let others = [];
  if (session.folderPath) {
    await flushMemos(session.id, session.folderPath).catch(() => {});
    try { others = await readAllMemos(session.folderPath); } catch { /* 端末内のメモだけで続行 */ }
  }
  return combine(mine, others);
}

// ===========================================================================
// 会議資料（REQ-10）
// ===========================================================================
const TEXT_TYPES = /^(text\/|application\/(json|xml|csv))/;
const INLINE_TYPES = /^(application\/pdf|image\/(png|jpeg|webp|heic|heif))$/;
const MAX_MATERIAL_BYTES = 40 * 1024 * 1024;

export async function addMaterial(session, file) {
  const rec = { id: uuid(), sessionId: session.id, name: safeFileName(file.name), type: file.type || guessType(file.name), size: file.size, blob: file, uploaded: false, addedAt: Date.now() };
  await putFile(rec);
  session.materials = [...(session.materials || []), { id: rec.id, name: rec.name, type: rec.type, size: rec.size }];
  await putSession(session);
  if (session.folderPath && getSettings().onedrive.clientId) await uploadPendingMaterials(session).catch(() => {});
  return rec;
}

function guessType(name) {
  const ext = String(name).toLowerCase().split('.').pop();
  return { pdf: 'application/pdf', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', txt: 'text/plain', md: 'text/markdown', csv: 'text/csv', json: 'application/json' }[ext] || 'application/octet-stream';
}

async function uploadPendingMaterials(session) {
  if (!session.folderPath) return;
  const files = await getFiles(session.id);
  for (const f of files) {
    if (f.uploaded) continue;
    await od.upload(od.joinPath(session.folderPath, '資料', f.name), f.blob, f.type);
    f.uploaded = true;
    await putFile(f);
  }
}

async function collectMaterials(session) {
  const out = [];
  let total = 0;
  const local = await getFiles(session.id);
  const names = new Set();
  const push = async (name, type, blob) => {
    if (names.has(name)) return;
    names.add(name);
    if (TEXT_TYPES.test(type) || /\.(txt|md|csv)$/i.test(name)) {
      out.push({ name, text: (await blob.text()).slice(0, 200000) });
    } else if (INLINE_TYPES.test(type)) {
      if (total + blob.size > MAX_MATERIAL_BYTES) { out.push({ name, text: '（容量が大きいため省略しました）' }); return; }
      total += blob.size;
      out.push({ name, mimeType: type, data: await blobToBase64(blob) });
    } else {
      out.push({ name, text: '（この形式は読み込めないため省略しました。PDF・画像・テキストに変換してください）' });
    }
  };
  for (const f of local) await push(f.name, f.type, f.blob);
  if (session.folderPath) {
    const remote = await od.listChildren(od.joinPath(session.folderPath, '資料')).catch(() => []);
    for (const r of remote) {
      if (r.folder || names.has(r.name)) continue;
      const res = await od.download(od.joinPath(session.folderPath, '資料', r.name), 'blob').catch(() => null);
      if (res) await push(r.name, (r.file && r.file.mimeType) || guessType(r.name), res.body);
    }
  }
  return out;
}

// ===========================================================================
// 会議の一覧（OneDrive上の _app/index.json）と、前回の持ち越しToDo（D-13）
// ===========================================================================
export async function readIndex() {
  try { return (await od.readJson(indexPath())) || { sessions: {} }; } catch { return { sessions: {} }; }
}

export async function updateIndex(session) {
  if (!session.folderPath) return;
  try {
    const idx = await readIndex();
    idx.sessions = idx.sessions || {};
    const m = session.minutes && session.minutes.data;
    idx.sessions[session.folderName] = {
      id: session.id, folderName: session.folderName, folderPath: session.folderPath, title: session.title,
      projectId: session.projectId, templateId: session.templateId, startedAt: session.startedAt, durationSec: session.durationSec,
      status: session.status, confidential: !!session.confidential,
      openTodos: m ? (m.todos || []).map((t) => ({ task: t.task, owner: t.owner, due: t.due })) : [],
      decisions: m ? (m.decisions || []).slice(0, 5).map((d) => d.text) : [],
      updatedAt: Date.now(),
    };
    await od.uploadJson(indexPath(), idx);
  } catch (e) { console.warn('一覧の更新に失敗', e); }
}

export async function previousTodos(session) {
  const prev = await previousSession(session);
  return prev ? prev.openTodos || [] : [];
}

/** 同じ案件（共通の場合は同じ会議名）の直前の会議 */
export async function previousSession(session) {
  const idx = await readIndex();
  const list = Object.values(idx.sessions || {})
    .filter((s) => s.id !== session.id && (s.startedAt || 0) < (session.startedAt || Date.now()))
    .filter((s) => (session.projectId && session.projectId !== vocab.COMMON_ID ? s.projectId === session.projectId : s.title === session.title))
    .sort((a, b) => (b.startedAt || 0) - (a.startedAt || 0));
  return list[0] || null;
}

/** OneDrive上の会議（他の端末で録音したもの）を読み込む */
export async function loadRemoteSession(folderPath) {
  const meta = await od.readJson(metaPath(folderPath, 'session.json'));
  if (!meta) throw new AppError('会議の情報が見つかりません。');
  const local = await getSession(meta.id);
  const s = { ...(local || {}), ...meta, folderPath, isRemote: !local };
  if (!s.transcript) s.transcript = await od.readJson(metaPath(folderPath, 'transcript.json')).catch(() => null);
  if (!s.minutes) s.minutes = await od.readJson(metaPath(folderPath, 'minutes.json')).catch(() => null);
  if (!s.levels) { const lv = await od.readJson(metaPath(folderPath, 'levels.json')).catch(() => null); if (lv) s.levels = lv.levels; }
  await putSession(s);
  return s;
}

/** 一覧：OneDriveの索引と端末内の会議を合わせる */
export async function listAll({ remote = true } = {}) {
  const local = await listSessions();
  const map = new Map(local.map((s) => [s.id, { ...s, where: 'local' }]));
  if (remote && getSettings().onedrive.clientId) {
    const idx = await readIndex();
    for (const e of Object.values(idx.sessions || {})) {
      const cur = map.get(e.id);
      if (cur) map.set(e.id, { ...e, ...cur, where: 'both' });
      else map.set(e.id, { ...e, where: 'remote' });
    }
  }
  return Array.from(map.values()).sort((a, b) => (b.startedAt || 0) - (a.startedAt || 0));
}

// ===========================================================================
// 手直しと学習（REQ-7）、話者名の一括変更（REQ-6 / C-10）
// ===========================================================================
export async function saveTranscriptEdits(session, edits) {
  const pairs = [];
  const segs = session.transcript.segments.map((s) => {
    if (edits[s.id] == null || edits[s.id] === s.text) return s;
    pairs.push({ before: s.text, after: edits[s.id] });
    return { ...s, text: edits[s.id], edited: true };
  });
  session.transcript = { ...session.transcript, segments: segs };
  await putSession(session);
  if (session.folderPath) await saveTranscriptFiles(session);
  let added = [];
  if (pairs.length) {
    try {
      const res = await ai.learnFromEdits(pairs, vocab.termsFor(session.projectId).map((t) => t.correct), { sessionId: session.id, kind: 'transcript' });
      added = await vocab.addTerms(session.projectId || vocab.COMMON_ID, res.terms || [], '手直し');
    } catch (e) { console.warn(e); }
  }
  return { changed: pairs.length, learned: added.length };
}

export async function saveMinutesEdits(session, markdown) {
  const before = session.minutes.markdown;
  session.minutes = { ...session.minutes, markdown, edited: true, editedAt: Date.now() };
  await putSession(session);
  if (session.folderPath) {
    await od.uploadText(od.joinPath(session.folderPath, '議事録.md'), markdown, 'text/markdown;charset=utf-8');
    await od.uploadJson(metaPath(session.folderPath, 'minutes.json'), session.minutes);
  }
  const hunks = tp.lineDiff(before, markdown);
  let learnedTerms = 0;
  let learnedRules = 0;
  if (hunks.length) {
    try {
      const pairs = hunks.map((h) => ({ before: h.before.join('\n') || '（なし）', after: h.after.join('\n') || '（削除）' }));
      const res = await ai.learnFromEdits(pairs, vocab.termsFor(session.projectId).map((t) => t.correct), { sessionId: session.id, kind: 'minutes' });
      learnedTerms = (await vocab.addTerms(session.projectId || vocab.COMMON_ID, res.terms || [], '手直し')).length;
      learnedRules = (await vocab.addRules(session.projectId || vocab.COMMON_ID, res.rules || [], '手直し')).length;
    } catch (e) { console.warn(e); }
  }
  return { learnedTerms, learnedRules };
}

export async function renameSpeakers(session, map) {
  const entries = Object.entries(map).filter(([a, b]) => b && b.trim() && a !== b.trim());
  if (!entries.length) return;
  const rename = (name) => { const hit = entries.find(([a]) => a === name); return hit ? hit[1].trim() : name; };
  session.transcript = { ...session.transcript, segments: session.transcript.segments.map((s) => ({ ...s, speaker: rename(s.speaker) })) };
  if (session.minutes && session.minutes.markdown) {
    let md = session.minutes.markdown;
    for (const [a, b] of entries) md = md.split(a).join(b.trim());
    session.minutes = { ...session.minutes, markdown: md };
  }
  await putSession(session);
  if (session.folderPath) {
    await saveTranscriptFiles(session);
    if (session.minutes && session.minutes.markdown) {
      await od.uploadText(od.joinPath(session.folderPath, '議事録.md'), session.minutes.markdown, 'text/markdown;charset=utf-8');
      await od.uploadJson(metaPath(session.folderPath, 'minutes.json'), session.minutes);
    }
  }
  await vocab.rememberSpeakers(session.projectId || vocab.COMMON_ID, entries.map(([, b]) => b));
}

// ===========================================================================
// 端末内の一時データの整理（機密保護）
// ===========================================================================
export async function cleanupLocal() {
  const days = Number(getSettings().privacy.localRetentionDays);
  if (!(days > 0)) return 0;
  const limit = Date.now() - days * 86400000;
  let n = 0;
  for (const s of await listSessions()) {
    if (s.steps && s.steps.audio && (s.audioSavedAt || s.endedAt || 0) < limit && s.status !== 'recording') {
      await deleteChunks(s.id);
      n++;
    }
  }
  return n;
}

/** 前回の録音が途中で終わっていた場合（ブラウザの終了など）の復元（A-4） */
export async function findInterrupted(activeId) {
  const out = [];
  for (const s of await listSessions()) {
    if (s.status === 'recording' && s.id !== activeId) {
      const chunks = await getChunks(s.id);
      const bytes = chunks.reduce((n, c) => n + c.byteLength, 0);
      s.durationSec = bytes / 2 / SAMPLE_RATE;
      s.endedAt = (s.startedAt || s.createdAt) + s.durationSec * 1000;
      s.status = bytes ? 'recorded' : 'error';
      s.recovered = true;
      if (!bytes) s.error = '録音データがありません。';
      await putSession(s);
      if (bytes) out.push(s);
    }
  }
  return out;
}
