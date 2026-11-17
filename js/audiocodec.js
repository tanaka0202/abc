// 録音データの圧縮（AAC・M4A形式）と、圧縮した音声を書き起こし用の音声（16kHz・モノラル）に戻す処理。
// 圧縮は、ブラウザ内蔵の機能（WebCodecs の AudioEncoder）を使う。端末によって使える設定が違うため、順に試して使えるものを選ぶ。
// 圧縮できない端末・圧縮の結果を確かめられなかった場合は、呼び出し側がWAV（無圧縮）で保存する。
import { buildM4a, readM4a, parseAdts, ascFor, AAC_FRAME } from './m4a.js';
import { AppError } from './util.js';

const PCM_RATE = 16000;

// 試す設定（上から順）。録音は16kHz・モノラルなので、そのまま圧縮できる設定を優先する。
// 16kHzを受け付けない端末（Windowsの一部など）では、48kHzに変換してから圧縮する（その分ファイルは大きくなる）。
const CANDIDATES = [
  { sampleRate: 16000, channels: 1, bitrate: 48000 },
  { sampleRate: 16000, channels: 1, bitrate: 64000 },
  { sampleRate: 16000, channels: 1, bitrate: 32000 },
  { sampleRate: 48000, channels: 1, bitrate: 64000 },
  { sampleRate: 48000, channels: 1, bitrate: 96000 },
  { sampleRate: 48000, channels: 1, bitrate: 128000 },
  { sampleRate: 48000, channels: 2, bitrate: 96000 },
  { sampleRate: 48000, channels: 2, bitrate: 128000 },
];
const failed = new Set(); // この起動中に失敗した設定は、もう試さない
const keyOf = (c) => `${c.sampleRate}/${c.channels}/${c.bitrate}`;
const encoderConfig = (c) => ({ codec: 'mp4a.40.2', sampleRate: c.sampleRate, numberOfChannels: c.channels, bitrate: c.bitrate, aac: { format: 'aac' } });

/** この端末で使える圧縮の設定を、優先順に並べて返す（使えなければ空） */
export async function supportedAacConfigs() {
  if (typeof AudioEncoder === 'undefined' || typeof AudioData === 'undefined') return [];
  const out = [];
  for (const c of CANDIDATES) {
    if (failed.has(keyOf(c))) continue;
    try { if ((await AudioEncoder.isConfigSupported(encoderConfig(c))).supported) out.push(c); } catch { /* 使えない */ }
  }
  return out;
}

// ---------------------------------------------------------------------------
// サンプルレートの変換（16kHz ⇔ 48kHz）。折り返し雑音が出ないよう、8kHzより上を切るフィルタを通す
// ---------------------------------------------------------------------------
const TAPS = 48;
const LOWPASS = (() => {
  // 48kHzで、7.4kHzより上を切る（窓関数つきsinc）
  const h = new Float32Array(TAPS);
  const fc = 7400 / 48000;
  const mid = (TAPS - 1) / 2;
  let sum = 0;
  for (let i = 0; i < TAPS; i++) {
    const x = i - mid;
    const sinc = x === 0 ? 2 * fc : Math.sin(2 * Math.PI * fc * x) / (Math.PI * x);
    const win = 0.5 - 0.5 * Math.cos((2 * Math.PI * (i + 0.5)) / TAPS);
    h[i] = sinc * win;
    sum += h[i];
  }
  for (let i = 0; i < TAPS; i++) h[i] /= sum;
  return h;
})();

/** 16kHz → 48kHz（3倍）。続けて呼ぶと、前回の続きとしてつながる */
export class Up3 {
  constructor() { this.hist = new Float32Array(TAPS / 3 - 1); }

  process(input) {
    const P = TAPS / 3;
    const x = new Float32Array(this.hist.length + input.length);
    x.set(this.hist);
    x.set(input, this.hist.length);
    const out = new Float32Array(input.length * 3);
    for (let n = 0; n < input.length; n++) {
      const base = n + this.hist.length;
      for (let p = 0; p < 3; p++) {
        let acc = 0;
        for (let k = 0; k < P; k++) acc += LOWPASS[p + 3 * k] * x[base - k];
        out[3 * n + p] = acc * 3;
      }
    }
    this.hist = x.slice(x.length - this.hist.length);
    return out;
  }
}

/** 48kHz → 16kHz（3分の1）。続けて呼ぶと、前回の続きとしてつながる */
export class Down3 {
  constructor() { this.hist = new Float32Array(TAPS - 1); this.phase = 0; }

  process(input) {
    const H = this.hist.length;
    const x = new Float32Array(H + input.length);
    x.set(this.hist);
    x.set(input, H);
    const out = new Float32Array(Math.ceil((input.length - this.phase) / 3));
    let m = 0;
    let i = this.phase;
    for (; i < input.length; i += 3) {
      let acc = 0;
      const base = i + H;
      for (let k = 0; k < TAPS; k++) acc += LOWPASS[k] * x[base - k];
      out[m++] = acc;
    }
    this.phase = i - input.length;
    this.hist = x.slice(x.length - H);
    return out.subarray(0, m);
  }
}

/** そのほかのサンプルレート → 16kHz（平均して間引く、簡易な方式） */
export class DownAny {
  constructor(rate) { this.ratio = rate / PCM_RATE; this.phase = 0; this.sum = 0; this.count = 0; this.last = 0; }

  process(input) {
    const out = new Float32Array(Math.ceil(input.length / this.ratio) + 2);
    let n = 0;
    for (let i = 0; i < input.length; i++) {
      this.sum += input[i]; this.count++;
      this.phase += 1;
      while (this.phase >= this.ratio) {
        this.phase -= this.ratio;
        this.last = this.count ? this.sum / this.count : this.last;
        out[n++] = this.last;
        this.sum = 0; this.count = 0;
      }
    }
    return out.subarray(0, n);
  }
}

const tick = () => new Promise((r) => setTimeout(r, 0));
/** 処理待ちがたまりすぎないよう、減るのを待つ（'dequeue' の知らせが来ない場合に備えて、時間でも抜ける） */
function waitDequeue(codec) {
  return new Promise((resolve) => {
    let done = false;
    const end = () => { if (!done) { done = true; resolve(); } };
    try { codec.addEventListener('dequeue', end, { once: true }); } catch { /* 無視 */ }
    setTimeout(end, 100);
  });
}

// ---------------------------------------------------------------------------
// 圧縮：PCM（16kHz・モノラル・16ビット）→ M4A
// ---------------------------------------------------------------------------
async function encodeWith(c, pcm, { onProgress, signal, createdAt }) {
  const frames = [];
  let desc = null;
  let adts = null;
  let failure = null;
  const enc = new AudioEncoder({
    output: (chunk, meta) => {
      try {
        const u = new Uint8Array(chunk.byteLength);
        chunk.copyTo(u);
        const d = meta && meta.decoderConfig && meta.decoderConfig.description;
        if (d && !desc) desc = new Uint8Array(d instanceof ArrayBuffer ? d.slice(0) : d.buffer.slice(d.byteOffset, d.byteOffset + d.byteLength));
        // 端末によっては、見出し（ADTS）つきで返すことがある。その場合は見出しを外して中身だけにする
        const a = !desc && u.length > 7 && u[0] === 0xff && (u[1] & 0xf6) === 0xf0 ? parseAdts(u) : null;
        if (a && a.frames.length === 1) { adts = a; frames.push(a.frames[0].slice()); } else frames.push(u);
      } catch (e) { failure = failure || e; }
    },
    error: (e) => { failure = failure || e; },
  });
  try {
    enc.configure(encoderConfig(c));
    const totalSamples = Math.floor(pcm.size / 2);
    const READ = PCM_RATE * 10; // 10秒ずつ読む
    const STEP = PCM_RATE;      // 1秒ずつ渡す
    const up = c.sampleRate === 48000 ? new Up3() : null;
    let fed = 0; // 圧縮する側のサンプルレートで数えた、渡したサンプル数
    for (let pos = 0; pos < totalSamples; pos += READ) {
      if (signal && signal.aborted) throw new AppError('中止しました。');
      const n = Math.min(READ, totalSamples - pos);
      const src = new Int16Array(await pcm.slice(pos * 2, (pos + n) * 2).arrayBuffer());
      for (let o = 0; o < n; o += STEP) {
        const m = Math.min(STEP, n - o);
        let f = new Float32Array(m);
        for (let i = 0; i < m; i++) f[i] = src[o + i] / 32768;
        if (up) f = up.process(f);
        let data = f;
        if (c.channels === 2) { data = new Float32Array(f.length * 2); data.set(f); data.set(f, f.length); } // 同じ音を左右に入れる
        const ad = new AudioData({ format: 'f32-planar', sampleRate: c.sampleRate, numberOfFrames: f.length, numberOfChannels: c.channels, timestamp: Math.round((fed / c.sampleRate) * 1e6), data });
        enc.encode(ad);
        ad.close();
        fed += f.length;
        if (failure) throw failure;
        while (enc.encodeQueueSize > 6) { await waitDequeue(enc); if (failure) throw failure; }
      }
      if (onProgress) onProgress(Math.min(0.99, (pos + n) / totalSamples));
      await tick(); // 画面が固まらないよう、ひと息入れる
    }
    await enc.flush();
    if (failure) throw failure;
  } finally {
    try { if (enc.state !== 'closed') enc.close(); } catch { /* 無視 */ }
  }
  if (!frames.length) throw new Error('圧縮の結果が空でした');
  const asc = desc || (adts && adts.asc) || ascFor(c.sampleRate, c.channels);
  const built = buildM4a({ frames, sampleRate: c.sampleRate, channels: c.channels, asc, bitrate: c.bitrate, createdAt });
  const blob = new Blob(built.parts, { type: 'audio/mp4' });
  return { blob, sampleRate: c.sampleRate, channels: c.channels, bitrate: c.bitrate, frames: built.frames, durationSec: built.durationSec };
}

/**
 * 録音データ（PCMのBlob）をM4Aに圧縮する。使える設定を順に試し、できあがったファイルを再生できることを確かめてから返す。
 * 圧縮できない・確かめられない場合は例外を投げる（呼び出し側はWAVで保存する）。
 */
export async function encodeM4a(pcm, { onProgress, signal, createdAt = 0 } = {}) {
  const configs = await supportedAacConfigs();
  if (!configs.length) throw new AppError('この端末のブラウザは、音声の圧縮（AAC）に対応していません');
  const expectSec = pcm.size / 2 / PCM_RATE;
  let lastErr = null;
  for (const c of configs) {
    try {
      const res = await encodeWith(c, pcm, { onProgress, signal, createdAt });
      await verifyM4a(res.blob, expectSec);
      if (onProgress) onProgress(1);
      return res;
    } catch (e) {
      if (signal && signal.aborted) throw e;
      failed.add(keyOf(c));
      lastErr = e;
    }
  }
  throw new AppError(`音声を圧縮できませんでした（${(lastErr && lastErr.message) || '原因不明'}）`, { cause: lastErr });
}

/** できあがったM4Aを確かめる：目次を読み直せること、長さが録音と合うこと、ブラウザが最後まで再生できる形であること */
export async function verifyM4a(blob, expectSec) {
  const info = await readM4a(blob);
  const tol = Math.max(0.6, expectSec * 0.005);
  if (Math.abs(info.durationSec - expectSec) > tol) throw new Error(`圧縮後の長さが合いません（録音 ${expectSec.toFixed(1)}秒／圧縮後 ${info.durationSec.toFixed(1)}秒）`);
  const last = info.offsets[info.count - 1] + info.sizes[info.count - 1];
  if (last > blob.size) throw new Error('圧縮後のファイルが途中で切れています');
  // (1) 最初と最後の部分を、実際に音声に戻せること
  const decoded = await spotDecode(blob, info);
  // (2) ブラウザの再生機能で開けて、最後の位置まで移動できること（確かめられない環境では (1) だけで判断する）
  let playable = null;
  try { playable = await playableCheck(blob, info.durationSec); } catch { playable = null; }
  if (playable === false) throw new Error('圧縮後のファイルを、ブラウザが再生できませんでした');
  if (!decoded && playable !== true) throw new Error('圧縮後のファイルを確かめられませんでした');
  return info;
}

/** ブラウザの再生機能で開けるか。true＝開けて、最後の位置まで移動できた／false＝開けない／null＝確かめられなかった */
function playableCheck(blob, durationSec) {
  if (typeof Audio === 'undefined' || typeof URL === 'undefined' || !URL.createObjectURL) return Promise.resolve(null);
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const a = new Audio();
    let stage = 0;
    let done = false;
    const end = (v) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      try { a.removeAttribute('src'); a.load(); } catch { /* 無視 */ }
      URL.revokeObjectURL(url);
      resolve(v);
    };
    const timer = setTimeout(() => end(null), 6000);
    a.muted = true;
    a.preload = 'auto';
    a.addEventListener('error', () => end(false));
    a.addEventListener('loadedmetadata', () => {
      if (!Number.isFinite(a.duration) || Math.abs(a.duration - durationSec) > Math.max(1, durationSec * 0.01)) { end(false); return; }
      stage = 1;
      try { a.currentTime = Math.max(0, a.duration - 0.4); } catch { end(null); }
    });
    // 移動した先の音声を取り出せたら合格（取り出せなければ error が届く）
    const ready = () => { if (stage === 2 && a.readyState >= 2) end(true); };
    a.addEventListener('seeked', () => { if (stage === 1) { stage = 2; ready(); } });
    a.addEventListener('canplay', ready);
    a.addEventListener('loadeddata', ready);
    a.src = url;
    try { a.load(); } catch { end(null); }
  });
}

/** 最初と最後の数フレームを、実際に音声に戻して確かめる。確かめる機能がないブラウザでは false */
async function spotDecode(blob, info) {
  if (typeof AudioDecoder === 'undefined' || typeof EncodedAudioChunk === 'undefined') return false;
  const pick = [];
  for (let i = 0; i < Math.min(4, info.count); i++) pick.push(i);
  for (let i = Math.max(4, info.count - 4); i < info.count; i++) pick.push(i);
  let got = 0;
  let failure = null;
  const dec = new AudioDecoder({ output: (d) => { got += d.numberOfFrames; d.close(); }, error: (e) => { failure = e; } });
  try {
    dec.configure({ codec: 'mp4a.40.2', sampleRate: info.sampleRate, numberOfChannels: info.channels, description: info.asc });
    for (const i of pick) {
      const data = await blob.slice(info.offsets[i], info.offsets[i] + info.sizes[i]).arrayBuffer();
      if (failure) break;
      dec.decode(new EncodedAudioChunk({ type: 'key', timestamp: Math.round(((i * info.frameSamples) / info.sampleRate) * 1e6), data }));
    }
    if (!failure) await dec.flush();
  } catch (e) {
    failure = failure || e; // 音声に戻せないと、途中で部品が閉じられて例外になることがある。先に届いた理由を優先する
  } finally { try { if (dec.state !== 'closed') dec.close(); } catch { /* 無視 */ } }
  if (failure) throw new Error(`圧縮後のファイルを音声に戻せませんでした（${failure.message || failure}）`);
  if (!got) throw new Error('圧縮後のファイルを音声に戻せませんでした');
  return true;
}

// ---------------------------------------------------------------------------
// 復元：M4A → PCM（16kHz・モノラル・16ビット）。OneDriveの録音データから書き起こすときに使う
// ---------------------------------------------------------------------------
/** M4A（Blob）を、少しずつ音声に戻してPCMのBlobにする（長い録音でもメモリを使いすぎない） */
export async function decodeM4aToPcm(blob, { onProgress, signal } = {}) {
  if (typeof AudioDecoder === 'undefined' || typeof EncodedAudioChunk === 'undefined') throw new AppError('このブラウザは、圧縮した録音データの読み取りに対応していません');
  const info = await readM4a(blob);
  const cfg = { codec: 'mp4a.40.2', sampleRate: info.sampleRate, numberOfChannels: info.channels, description: info.asc };
  const sup = await AudioDecoder.isConfigSupported(cfg).catch(() => ({ supported: false }));
  if (!sup.supported) throw new AppError('このブラウザは、この録音データ（AAC）を読み取れません');
  const parts = [];
  let failure = null;
  let conv = null;
  let convRate = 0;
  const dec = new AudioDecoder({
    output: (d) => {
      try {
        const n = d.numberOfFrames;
        const ch = d.numberOfChannels;
        let f = new Float32Array(n);
        d.copyTo(f, { planeIndex: 0, format: 'f32-planar' });
        if (ch > 1) {
          const g = new Float32Array(n);
          for (let c = 1; c < ch; c++) { d.copyTo(g, { planeIndex: c, format: 'f32-planar' }); for (let i = 0; i < n; i++) f[i] += g[i]; }
          for (let i = 0; i < n; i++) f[i] /= ch;
        }
        const rate = d.sampleRate;
        d.close();
        if (rate !== PCM_RATE) {
          if (!conv || convRate !== rate) { conv = rate === 48000 ? new Down3() : new DownAny(rate); convRate = rate; }
          f = conv.process(f);
        }
        const out = new Int16Array(f.length);
        for (let i = 0; i < f.length; i++) { const x = Math.max(-1, Math.min(1, f[i])); out[i] = x < 0 ? Math.round(x * 32768) : Math.round(x * 32767); }
        parts.push(out);
      } catch (e) { failure = failure || e; try { d.close(); } catch { /* 無視 */ } }
    },
    error: (e) => { failure = failure || e; },
  });
  try {
    dec.configure(cfg);
    const BATCH = 256; // まとめて読むフレーム数
    for (let i = 0; i < info.count; i += BATCH) {
      if (signal && signal.aborted) throw new AppError('中止しました。');
      const last = Math.min(info.count, i + BATCH) - 1;
      const start = info.offsets[i];
      const buf = new Uint8Array(await blob.slice(start, info.offsets[last] + info.sizes[last]).arrayBuffer());
      for (let k = i; k <= last; k++) {
        const o = info.offsets[k] - start;
        dec.decode(new EncodedAudioChunk({ type: 'key', timestamp: Math.round(((k * info.frameSamples) / info.sampleRate) * 1e6), duration: Math.round((info.frameSamples / info.sampleRate) * 1e6), data: buf.subarray(o, o + info.sizes[k]) }));
      }
      if (failure) throw failure;
      while (dec.decodeQueueSize > BATCH) { await waitDequeue(dec); if (failure) throw failure; }
      if (onProgress) onProgress(Math.min(0.99, (last + 1) / info.count));
      await tick();
    }
    await dec.flush();
    if (failure) throw failure;
  } finally { try { if (dec.state !== 'closed') dec.close(); } catch { /* 無視 */ } }
  if (!parts.length) throw new AppError('録音データを音声に戻せませんでした');
  if (onProgress) onProgress(1);
  return new Blob(parts, { type: 'application/octet-stream' });
}

export { AAC_FRAME };
