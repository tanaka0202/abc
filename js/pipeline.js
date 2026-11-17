// 処理の流れ：①録音保存＆書き起こし作成、②メモを組み込んで議事録作成＆保存（REQ-5）
// 各手順の完了を session.steps に記録し、失敗しても「再試行」で続きから再開できる。
import * as od from './onedrive.js';
import * as ai from './gemini.js';
import { upsertEvent } from './gcal.js';
import { googleEnsure } from './auth.js';
import { getChunks, putSession, getSession, listSessions, getFiles, putFile, deleteChunks, deleteFile, deleteSessionRecord, addChunk } from './db.js';
import { getSettings } from './settings.js';
import { rootPath, metaPath, readAllMemos } from './sync.js';
import * as tp from './textproc.js';
import * as vocab from './vocab.js';
import { getMyMemos, flushMemos, combine, clearMyMemos } from './memos.js';
import { pcmBlobFromChunks, wavFromPcmBlob, wavSlice, parseWav, SAMPLE_RATE } from './wav.js';
import { encodeM4a, decodeM4aToPcm } from './audiocodec.js';
import { logKaizen } from './usage.js';
import { AppError, folderNameFor, fmtDateTime, fmtTime, blobToBase64, safeFileName, uuid, errorMessage, isMobile } from './util.js';

const indexPath = () => od.joinPath(rootPath(), '_app', 'index.json');

/** OneDriveの会議フォルダを用意する */
export async function ensureFolder(session) {
  if (session.folderPath) return session;
  const created = await od.createUniqueFolder(rootPath(), folderNameFor(session.startedAt || session.createdAt || Date.now()));
  session.folderPath = created.path;
  session.folderName = created.name;
  if (created.webUrl) session.folderWebUrl = created.webUrl;
  else { try { session.folderWebUrl = await od.webUrl(created.path); } catch { /* 後で取得 */ } }
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
    diarization: s.diarization !== false, scene: s.scene || null, imported: !!s.imported, zeroSec: s.zeroSec || 0,
    audioFile: s.audioFile || null, audioInfo: s.audioInfo || null,
    processingBy: s.processingBy || null, processingName: s.processingName || null,
    reservedBy: s.reservedBy || null, reservedAt: s.reservedAt || null,
    materials: (s.materials || []).map((m) => ({ id: m.id, name: m.name, type: m.type, size: m.size })),
    updatedAt: Date.now(),
  };
}

export async function writeMeta(session) {
  if (!session.folderPath) return;
  try { await od.uploadJson(metaPath(session.folderPath, 'session.json'), sessionMeta(session)); } catch (e) { console.warn(e); }
}

let lastMetaWrite = 0;
const sideNotes = new Map(); // 会議ID → 同時に進めている作業の一言（「録音データの保存 45%」など）
const lastMain = new Map();  // 会議ID → 直近の進み具合（一言だけ変わったときに、表示を出し直すため）
async function progress(session, onProgress, step, text, ratio = null) {
  lastMain.set(session.id, { step, text, ratio, onProgress });
  const side = sideNotes.get(session.id);
  session.progress = { step, text: side ? `${text}／${side}` : text, ratio, at: Date.now() };
  if (onProgress) onProgress(session.progress);
  if (Date.now() - lastMetaWrite > 8000 || ratio === 1) { lastMetaWrite = Date.now(); writeMeta(session); }
}
/** 同時に進めている作業の一言を変える（空文字で消す） */
function setSideNote(session, note) {
  if (note) sideNotes.set(session.id, note); else sideNotes.delete(session.id);
  const m = lastMain.get(session.id);
  if (m) progress(session, m.onProgress, m.step, m.text, m.ratio);
}

export const AUDIO_M4A = '録音データ.m4a';
export const AUDIO_WAV = '録音データ.wav';

/** OneDrive上の録音データのファイル名の候補（記録があればそれを先に。古い会議はWAV） */
function audioNames(session) {
  return [...new Set([session.audioFile, AUDIO_M4A, AUDIO_WAV].filter(Boolean))];
}

/** OneDrive上の録音データを探す。戻り値：{ name, item } または null */
export async function findAudio(session) {
  if (!session.folderPath) return null;
  for (const name of audioNames(session)) {
    const item = await od.getItem(od.joinPath(session.folderPath, name));
    if (item && !item.folder) return { name, item };
  }
  return null;
}

/** 再生用の一時的なURL（約1時間有効） */
export async function audioUrl(session) {
  const hit = await findAudio(session);
  return hit ? hit.item['@microsoft.graph.downloadUrl'] || null : null;
}

/** 録音データ（PCM）を取得：端末内にあればそれを、なければOneDriveの録音データ（M4AまたはWAV）を音声に戻して使う */
async function getPcm(session, onProgress) {
  const chunks = await getChunks(session.id);
  if (chunks.length) return pcmBlobFromChunks(chunks);
  if (!session.folderPath) throw new AppError('録音データが見つかりません。');
  const hit = await findAudio(session);
  if (!hit) throw new AppError('OneDriveに録音データが見つかりません。');
  await progress(session, onProgress, 'audio', 'OneDriveから録音データを読み込んでいます');
  const res = await od.download(od.joinPath(session.folderPath, hit.name), 'blob');
  if (!res) throw new AppError('OneDriveに録音データが見つかりません。');
  if (/\.m4a$/i.test(hit.name)) {
    try {
      return await decodeM4aToPcm(res.body);
    } catch (e) {
      // 少しずつ戻す方法が使えないブラウザでは、まとめて戻す（メモリを多く使う）
      const next = await decodedBlockReader(res.body);
      const parts = [];
      for (;;) {
        const block = await next();
        if (!block) break;
        const out = new Int16Array(block.length);
        for (let i = 0; i < block.length; i++) { const x = Math.max(-1, Math.min(1, block[i])); out[i] = x < 0 ? Math.round(x * 32768) : Math.round(x * 32767); }
        parts.push(out);
      }
      if (!parts.length) throw e;
      return new Blob(parts, { type: 'application/octet-stream' });
    }
  }
  const { pcm, sampleRate } = await parseWav(res.body);
  if (sampleRate !== SAMPLE_RATE) throw new AppError('このアプリで録音したWAV（16kHz）以外は再処理できません。');
  return pcm;
}

/**
 * OneDriveに保存する録音データを用意する。設定が「圧縮する」なら、M4A（AAC）に圧縮する。
 * 圧縮できない端末・圧縮の結果を確かめられなかった場合は、これまでどおりWAV（無圧縮）にする（録音を失わないことを優先）。
 */
async function prepareAudioFile(session, pcm, { onNote, signal }) {
  const wav = () => ({ name: AUDIO_WAV, blob: wavFromPcmBlob(pcm), type: 'audio/wav', info: { format: 'wav', bytes: pcm.size + 44 } });
  if ((getSettings().recording.audioFormat || 'm4a') !== 'm4a') return wav();
  try {
    const res = await encodeM4a(pcm, { signal, createdAt: session.startedAt || 0, onProgress: onNote });
    return {
      name: AUDIO_M4A, blob: res.blob, type: 'audio/mp4',
      info: { format: 'm4a', bytes: res.blob.size, sampleRate: res.sampleRate, channels: res.channels, bitrate: res.bitrate, wavBytes: pcm.size + 44 },
    };
  } catch (e) {
    if (signal && signal.aborted) throw e;
    const f = wav();
    f.info.note = `圧縮できなかったため、WAV（無圧縮）で保存しました：${errorMessage(e)}`;
    await logKaizen('fallback', `「${session.title}」の録音データを圧縮できず、WAVで保存（${errorMessage(e)}）`).catch(() => {});
    return f;
  }
}

// ===========================================================================
// ① 録音保存 ＆ 書き起こし作成
// ===========================================================================
/**
 * phase：'all'＝保存と書き起こし／'save'＝OneDriveへの保存まで（書き起こしは別の端末か、後で行う）。
 * 保存済みの会議に対して呼ぶと、書き起こしだけを行う（録音データは、端末内になければOneDriveから取得する）。
 */
export async function runStep1(session, { onProgress, signal, phase = 'all' } = {}) {
  const st0 = getSettings();
  session.status = 'processing';
  session.error = null;
  session.steps = session.steps || {};
  session.processingBy = st0.deviceId;
  session.processingName = st0.deviceName;
  await putSession(session);
  try {
    // 1. 保存先フォルダ
    await progress(session, onProgress, 'folder', 'OneDriveに保存先フォルダを用意しています');
    await ensureFolder(session);
    await writeMeta(session);

    const pcm = await getPcm(session, onProgress);
    if (!session.durationSec) session.durationSec = pcm.size / 2 / SAMPLE_RATE;
    if (!session.endedAt && session.startedAt) session.endedAt = session.startedAt + session.durationSec * 1000;

    const saveOnly = phase === 'save' && !session.steps.transcript;
    const needTranscript = !saveOnly && !session.steps.transcript;

    // 2. 録音データの保存（書き起こしも行う場合は、待たずに並行して進める）
    const saveAudio = async () => {
      if (session.steps.audio) return;
      // 進み具合：圧縮は全体の前半、保存は後半として表示する
      const note = (what, r) => {
        const pct = Math.round(r * 100);
        if (needTranscript) setSideNote(session, r < 1 || what === '圧縮' ? `録音データの${what} ${pct}%` : '');
        else progress(session, onProgress, 'audio', `録音データを${what === '圧縮' ? '圧縮' : '保存'}しています（${pct}%）`, what === '圧縮' ? r * 0.5 : 0.5 + r * 0.5);
      };
      const report = (r) => note('保存', r);
      const file = await prepareAudioFile(session, pcm, { signal, onNote: (r) => note('圧縮', r) });
      report(0);
      await Promise.all([
        od.upload(od.joinPath(session.folderPath, file.name), file.blob, file.type, report),
        od.uploadJson(metaPath(session.folderPath, 'levels.json'), { levels: session.levels || [], gaps: session.gaps || [] }),
      ]);
      session.audioFile = file.name;
      session.audioInfo = file.info;
      session.steps.audio = true;
      session.audioSavedAt = Date.now();
      await putSession(session);
      if (needTranscript) setSideNote(session, '');
    };
    const saveExtras = async () => {
      await flushMemos(session.id, session.folderPath).catch(() => {});
      await uploadPendingMaterials(session);
    };

    // 保存だけの場合：ここで終える（書き起こしは、別の端末か、後でこの端末で行う）
    if (saveOnly) {
      if (!session.steps.audio) await progress(session, onProgress, 'audio', '録音データを保存しています', 0);
      await Promise.all([saveAudio(), saveExtras()]);
      await progress(session, onProgress, 'calendar', 'カレンダーに登録しています');
      await tryCalendar(session).catch((e) => { session.calendarError = errorMessage(e); });
      session.status = 'saved';
      session.progress = null;
      session.processingBy = null;
      session.processingName = null;
      await putSession(session);
      await Promise.all([writeMeta(session), updateIndex(session)]);
      await logKaizen('session', `保存完了：「${session.title}」（${fmtTime(session.durationSec)}）`);
      return session;
    }

    // 3. 書き起こし（録音データの保存と同時に進める。片方が失敗しても、済んだ方は済みとして残し、再試行では残りだけ行う）
    if (needTranscript) await progress(session, onProgress, 'transcribe', '書き起こしを始めています', 0);
    else if (!session.steps.audio) await progress(session, onProgress, 'audio', '録音データを保存しています', 0);
    const [aRes, tRes, xRes] = await Promise.allSettled([
      saveAudio(),
      needTranscript ? transcribe(session, pcm, { onProgress, signal }) : Promise.resolve(),
      saveExtras(),
    ]);
    setSideNote(session, '');
    if (tRes.status === 'fulfilled' && needTranscript) {
      session.steps.transcript = true;
      await putSession(session);
    }
    for (const r of [tRes, aRes, xRes]) if (r.status === 'rejected') throw r.reason;

    // 4. カレンダー（サインインが有効な場合のみ。切れていれば②で登録）
    await progress(session, onProgress, 'calendar', 'カレンダーに登録しています');
    await tryCalendar(session).catch((e) => { session.calendarError = errorMessage(e); });

    session.status = session.steps.minutes ? 'done' : 'transcribed';
    session.progress = null;
    session.processingBy = null;
    session.processingName = null;
    await putSession(session);
    await Promise.all([writeMeta(session), updateIndex(session)]);
    await logKaizen('session', `①完了：「${session.title}」（${fmtTime(session.durationSec)}、発言${(session.transcript?.segments || []).length}件）`);
    return session;
  } catch (e) {
    sideNotes.delete(session.id);
    session.status = 'error';
    session.error = errorMessage(e);
    session.processingBy = null;
    session.processingName = null;
    await putSession(session);
    await writeMeta(session);
    if (session.steps.audio) await updateIndex(session);
    throw e;
  } finally {
    sideNotes.delete(session.id);
    lastMain.delete(session.id);
    if (onProgress) onProgress(null);
  }
}

/** 処理の進み具合の順位（端末内とOneDriveのどちらが新しいかの判定に使う） */
export function statusRank(s) {
  const st = s && s.steps ? s.steps : {};
  if (!s) return -1;
  if (st.minutes || s.status === 'done') return 4;
  if (st.transcript || s.status === 'transcribed') return 3;
  if (s.status === 'saved' || st.audio) return 2;
  return 1;
}

/**
 * ほかの端末が保存した、書き起こし待ちの会議を探す（PCが自動で書き起こすため）。
 * 一覧（_app/index.json）で候補を絞り、各会議の状態（_meta/session.json）で、だれも処理していないことを確かめる。
 */
export async function findWaitingTranscription({ max = 3, staleMs = 10 * 60 * 1000 } = {}) {
  const idx = await readIndex();
  const cands = Object.values(idx.sessions || {})
    .filter((e) => e.status === 'saved' || e.status === 'processing')
    .sort((a, b) => (a.startedAt || 0) - (b.startedAt || 0));
  const out = [];
  for (const e of cands) {
    if (out.length >= max) break;
    const meta = await od.readJson(metaPath(e.folderPath, 'session.json')).catch(() => null);
    if (!meta || !(meta.steps && meta.steps.audio) || (meta.steps && meta.steps.transcript)) continue;
    const busy = meta.status === 'processing' && Date.now() - (meta.updatedAt || 0) < staleMs;
    if (meta.status === 'saved' || (meta.status === 'processing' && !busy)) out.push(meta);
  }
  return out;
}

// ===========================================================================
// 音声ファイルの取り込み（端末の録音アプリやICレコーダーで録った音声）
// ===========================================================================
const IMPORT_FLUSH = SAMPLE_RATE * 5;

/** 16kHz・モノラルの音声（Float32Array を順に返す関数）を、録音と同じ形で端末内に保存する */
async function storeImported(session, nextBlock) {
  let seq = 0;
  let total = 0;
  let buf = new Int16Array(IMPORT_FLUSH);
  let len = 0;
  const levels = [];
  let secSum = 0;
  let secCount = 0;
  let zero = 0;
  const flush = async () => {
    if (!len) return;
    await addChunk(session.id, seq++, buf.slice(0, len).buffer);
    len = 0;
  };
  for (;;) {
    const block = await nextBlock();
    if (!block) break;
    for (let i = 0; i < block.length; i++) {
      const x = Math.max(-1, Math.min(1, block[i]));
      const v = x < 0 ? Math.round(x * 32768) : Math.round(x * 32767);
      buf[len++] = v;
      total++;
      if (v === 0) zero++;
      secSum += x * x;
      if (++secCount >= SAMPLE_RATE) {
        const rms = Math.sqrt(secSum / secCount);
        levels.push(Math.round((rms > 0 ? Math.max(-100, 20 * Math.log10(rms)) : -100) * 10) / 10);
        secSum = 0; secCount = 0;
      }
      if (len === IMPORT_FLUSH) await flush();
    }
  }
  await flush();
  return { total, levels, chunkCount: seq, zeroSec: Math.round(zero / SAMPLE_RATE) };
}

/** 16ビットPCMのWAVを、少しずつ読みながら16kHz・モノラルに変換する（長い録音でもメモリを使いすぎない） */
function wavBlockReader(file, info) {
  const frame = info.channels * 2;
  const ratio = info.sampleRate / SAMPLE_RATE;
  const STEP = 1 << 20; // 約1MBずつ
  let pos = info.dataOffset;
  const end = info.dataOffset + Math.floor(info.dataBytes / frame) * frame;
  let phase = 0;
  let sum = 0;
  let count = 0;
  let last = 0;
  return async () => {
    if (pos >= end) return null;
    const take = Math.min(end - pos, Math.floor(STEP / frame) * frame);
    const dv = new DataView(await file.slice(pos, pos + take).arrayBuffer());
    pos += take;
    const frames = take / frame;
    const out = new Float32Array(Math.ceil(frames / ratio) + 2);
    let n = 0;
    for (let i = 0; i < frames; i++) {
      let x = 0;
      for (let c = 0; c < info.channels; c++) x += dv.getInt16(i * frame + c * 2, true);
      x = x / info.channels / 32768;
      sum += x; count++;
      phase += 1;
      while (phase >= ratio) {
        phase -= ratio;
        last = count ? sum / count : last;
        out[n++] = last;
        sum = 0; count = 0;
      }
    }
    return out.subarray(0, n);
  };
}

async function readWavInfo(file) {
  const dv = new DataView(await file.slice(0, Math.min(file.size, 1 << 16)).arrayBuffer());
  const tag = (o) => String.fromCharCode(dv.getUint8(o), dv.getUint8(o + 1), dv.getUint8(o + 2), dv.getUint8(o + 3));
  if (dv.byteLength < 44 || tag(0) !== 'RIFF' || tag(8) !== 'WAVE') return null;
  let o = 12;
  let fmt = null;
  while (o + 8 <= dv.byteLength) {
    const id = tag(o);
    const size = dv.getUint32(o + 4, true);
    if (id === 'fmt ') fmt = { format: dv.getUint16(o + 8, true), channels: dv.getUint16(o + 10, true), sampleRate: dv.getUint32(o + 12, true), bits: dv.getUint16(o + 22, true) };
    if (id === 'data') {
      if (!fmt || fmt.format !== 1 || fmt.bits !== 16 || !fmt.channels || !fmt.sampleRate) return null;
      return { ...fmt, dataOffset: o + 8, dataBytes: Math.min(size, file.size - (o + 8)) };
    }
    o += 8 + size + (size % 2);
  }
  return null;
}

/** WAV以外（m4a・mp3など）：ブラウザの機能で16kHzに変換する（全体をメモリに展開するため、長い録音はPC向け） */
async function decodedBlockReader(file) {
  const Offline = window.OfflineAudioContext || window.webkitOfflineAudioContext;
  if (!Offline) throw new AppError('このブラウザでは、この形式の音声を取り込めません。WAV形式に変換してください。');
  let audio;
  try {
    audio = await new Offline(1, 1, SAMPLE_RATE).decodeAudioData(await file.arrayBuffer());
  } catch (e) {
    throw new AppError('この音声ファイルを読み取れませんでした（対応していない形式か、ファイルが大きすぎます）。WAV・m4a・mp3で試すか、PCで取り込んでください。', { cause: e });
  }
  const chs = [];
  for (let c = 0; c < audio.numberOfChannels; c++) chs.push(audio.getChannelData(c));
  const BLOCK = SAMPLE_RATE * 10;
  let pos = 0;
  return async () => {
    if (pos >= audio.length) return null;
    const n = Math.min(BLOCK, audio.length - pos);
    const out = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      let x = 0;
      for (const ch of chs) x += ch[pos + i];
      out[i] = x / chs.length;
    }
    pos += n;
    return out;
  };
}

/**
 * 音声ファイルを取り込み、「録音済み（未保存）」の会議として端末内に置く。以後は、このアプリの録音と同じように保存・書き起こしできる。
 * info：{ id?, title, projectId, templateId, participants, confidential, diarization, startedAt }
 */
export async function importAudioFile(file, info) {
  if (!file || !file.size) throw new AppError('音声ファイルが空です。');
  const st = getSettings();
  const session = {
    id: info.id || uuid(), title: (info.title || '').trim() || `会議 ${fmtDateTime(info.startedAt || Date.now())}`,
    projectId: info.projectId || vocab.COMMON_ID, templateId: info.templateId || 'standard', participants: info.participants || [],
    confidential: !!info.confidential, diarization: info.diarization !== false,
    source: 'import', scene: 'import', imported: true, importedName: file.name || '',
    status: 'importing', createdAt: Date.now(), steps: {}, deviceId: st.deviceId, deviceName: st.deviceName, materials: [],
  };
  const proj = vocab.projectById(session.projectId);
  session.projectName = proj ? proj.name : '';
  await deleteChunks(session.id);
  const wav = await readWavInfo(file);
  const reader = wav ? wavBlockReader(file, wav) : await decodedBlockReader(file);
  const res = await storeImported(session, reader);
  if (!res.total) { await deleteChunks(session.id); throw new AppError('音声ファイルに音声が含まれていません。'); }
  session.durationSec = res.total / SAMPLE_RATE;
  session.levels = res.levels;
  session.gaps = [];
  session.chunkCount = res.chunkCount;
  session.zeroSec = res.zeroSec;
  // 開始時刻：指定がなければ、ファイルの更新時刻（多くの録音アプリでは録音の終了時刻）から長さを引いて推定する
  session.startedAt = info.startedAt || Math.round((file.lastModified || Date.now()) - session.durationSec * 1000);
  session.endedAt = session.startedAt + session.durationSec * 1000;
  session.status = 'recorded';
  await putSession(session);
  return session;
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

/** 書き起こしを同時に進める区切りの数：PCは3、スマホ・タブレットは2（メモリが少ない端末は1） */
export function parallelSlices() {
  const set = Number(getSettings().gemini.parallel);
  if (set >= 1) return Math.min(4, Math.floor(set));
  if (!isMobile()) return 3;
  const mem = typeof navigator !== 'undefined' ? Number(navigator.deviceMemory) || 0 : 0;
  return mem && mem <= 4 ? 1 : 2;
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
  // 区切りを同時に複数送って、待ち時間を縮める（結果は元の順に並べる）。スマホはメモリを考えて控えめにする
  const conc = Math.min(plan.length, parallelSlices());
  const results = new Array(plan.length);
  let next = 0;
  let done = 0;
  let failed = false;
  const worker = async () => {
    for (;;) {
      const i = next++;
      if (i >= plan.length || failed) return;
      if (signal && signal.aborted) throw new AppError('中止しました。');
      const c = plan[i];
      const opts = {
        chunkIndex: i, offsetSec: c.sliceStart, durSec: c.sliceEnd - c.sliceStart, diarization,
        vocab: terms.map((t) => t.correct), participants: session.participants, sessionId: session.id, signal,
      };
      const silent = tp.isSilent(session.levels, c.sliceStart, c.sliceEnd);
      try {
        const segs = await transcribeSlice(wavSlice(pcm, c.sliceStart, c.sliceEnd), opts, { silent, state, session });
        results[i] = { ...c, segs };
      } catch (e) {
        failed = true;
        throw e;
      }
      done++;
      if (!failed) await progress(session, onProgress, 'transcribe', `書き起こしています（${done}/${plan.length}）`, (done / plan.length) * 0.88);
    }
  };
  await progress(session, onProgress, 'transcribe', `書き起こしています（0/${plan.length}）`, 0);
  await Promise.all(Array.from({ length: conc }, worker));
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
  await Promise.all([
    od.uploadText(od.joinPath(session.folderPath, '書き起こし.txt'), tp.segmentsToTxt(segs, meta)),
    od.uploadText(od.joinPath(session.folderPath, '字幕.srt'), tp.segmentsToSrt(segs)),
    od.uploadJson(metaPath(session.folderPath, 'transcript.json'), session.transcript),
  ]);
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
  // サインインが切れていたら、画面を出さずにつなぎ直す。できなければ保留にして、つながったときに登録する
  if (!(await googleEnsure(st.google.clientId, st.google.account).catch(() => false))) { session.calendarPending = true; await putSession(session); return false; }
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

    // メモ・資料・前回のToDoは、同時に集める
    await progress(session, onProgress, 'minutes', 'メモと会議資料を読み込んでいます', 0.15);
    const [memos, materials, carry] = await Promise.all([collectMemos(session), collectMaterials(session), previousTodos(session)]);
    const memoLines = memos.map((m) => `[${Number.isFinite(m.offset) ? (m.offset < 0 ? '開始前' : fmtTime(m.offset)) : '時刻なし'}] ${m.kind === 'star' ? '★' : ''}${m.text}`);

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

    // 保存・カレンダーの更新・資料からの用語の学習（REQ-7）は、同時に進める
    await progress(session, onProgress, 'save', '議事録・メモを保存しています', 0.85);
    const newMaterials = materials.filter((m) => !(session.learnedMaterials || []).includes(m.name));
    const learn = async () => {
      if (!newMaterials.length) return;
      try {
        const existing = vocab.termsFor(session.projectId).map((t) => t.correct);
        const res = await ai.extractTermsFromMaterials(newMaterials, existing, { sessionId: session.id });
        const added = await vocab.addTerms(session.projectId || vocab.COMMON_ID, res.terms || [], '資料');
        session.learnedMaterials = [...(session.learnedMaterials || []), ...newMaterials.map((m) => m.name)];
        session.learnedTerms = (session.learnedTerms || 0) + added.length;
      } catch (e) { session.learnError = errorMessage(e); }
    };
    const [saved] = await Promise.allSettled([
      saveMinutesFiles(session, memos),
      tryCalendar(session, minutes).catch((e) => { session.calendarError = errorMessage(e); }),
      learn(),
    ]);
    if (saved.status === 'rejected') throw saved.reason;

    session.steps = { ...(session.steps || {}), minutes: true };
    session.status = 'done';
    session.progress = null;
    await putSession(session);
    await Promise.all([writeMeta(session), updateIndex(session)]);
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
  await Promise.all([
    od.uploadText(od.joinPath(fp, '議事録.md'), session.minutes.markdown, 'text/markdown;charset=utf-8'),
    memos ? od.uploadText(od.joinPath(fp, 'メモ.txt'), tp.memosToTxt(memos, meta)) : null,
    od.uploadText(od.joinPath(fp, 'ToDo一覧.csv'), tp.todosToCsv(session.minutes.data.todos, meta), 'text/csv;charset=utf-8'),
    od.uploadText(od.joinPath(fp, 'カレンダー登録.ics'), tp.buildIcs(calendarEvent(session, session.minutes.data)), 'text/calendar;charset=utf-8'),
    od.uploadJson(metaPath(fp, 'minutes.json'), session.minutes),
  ]);
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
  // ほかの端末で書き起こし・議事録が作られていたら、端末内の古い写しではなくOneDriveのものを読む
  if (local && statusRank(meta) > statusRank(local)) { s.transcript = null; s.minutes = null; }
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
      // ほかの端末で処理が進んでいる場合（PCが書き起こした、など）は、OneDrive側の状態を表示する
      if (cur && statusRank(e) > statusRank(cur)) map.set(e.id, { ...cur, ...e, where: 'both', remoteNewer: true });
      else if (cur) map.set(e.id, { ...e, ...cur, where: 'both' });
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

/**
 * 保存前の録音を破棄する（不要な録音を残さないため）。
 * 端末内の録音・メモ・資料を消し、録音開始時にOneDriveへ作ったフォルダも、音声などをまだ保存していなければ消す。
 * 戻り値：{ folder: 'none'|'removed'|'kept'|'failed' }
 */
export async function discardSession(session) {
  if (session.steps && session.steps.audio) throw new AppError('この会議はOneDriveに保存済みのため、ここからは破棄できません。');
  await deleteChunks(session.id);
  for (const f of await getFiles(session.id)) await deleteFile(f.id);
  await clearMyMemos(session.id);
  await deleteSessionRecord(session.id);
  if (!session.folderPath) return { folder: 'none' };
  try {
    const kids = await od.listChildren(session.folderPath);
    // 録音開始時に作るのは _meta（進行状況・メモ）だけ。ほかのファイルがあれば、消さずに残す
    if (kids.some((k) => k.name !== '_meta')) return { folder: 'kept' };
    await od.remove(session.folderPath);
    return { folder: 'removed' };
  } catch {
    return { folder: 'failed' };
  }
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
