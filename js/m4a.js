// M4A（AACの音声を入れたMP4ファイル）の組み立てと読み取り。外部ライブラリは使わない。
// 録音データを小さくして保存するために使う。スマホ・Windows・Macの標準のプレーヤーや、OneDriveの画面でそのまま再生できる形式。
// ファイルの並びは「ftyp → moov（目次）→ mdat（音声）」。目次を先頭に置くので、再生をすぐ始められ、好きな位置に移動できる。

export const AAC_FRAME = 1024; // AAC-LC の1フレームのサンプル数
const SAMPLE_RATES = [96000, 88200, 64000, 48000, 44100, 32000, 24000, 22050, 16000, 12000, 11025, 8000, 7350];

const str = (s) => Uint8Array.from(s, (c) => c.charCodeAt(0));
const u8 = (...a) => Uint8Array.from(a);
const u16 = (n) => u8((n >>> 8) & 255, n & 255);
const u32 = (n) => u8((n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255);
const zeros = (n) => new Uint8Array(n);

function concat(parts) {
  let len = 0;
  for (const p of parts) len += p.length;
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

function box(type, ...payload) {
  const body = concat(payload);
  return concat([u32(body.length + 8), str(type), body]);
}
const fullBox = (type, version, flags, ...payload) => box(type, u8(version, (flags >>> 16) & 255, (flags >>> 8) & 255, flags & 255), ...payload);

/** MPEG-4の「記述子」（タグ＋長さ＋中身）。長さは7ビットずつの可変長 */
function descriptor(tag, ...payload) {
  const body = concat(payload);
  const len = [];
  let n = body.length;
  do { len.unshift(n & 0x7f); n >>>= 7; } while (n > 0);
  for (let i = 0; i < len.length - 1; i++) len[i] |= 0x80;
  return concat([u8(tag, ...len), body]);
}

/** AACの設定情報（AudioSpecificConfig）：AAC-LC・サンプルレート・チャンネル数から2バイトを作る */
export function ascFor(sampleRate, channels) {
  const idx = SAMPLE_RATES.indexOf(sampleRate);
  if (idx < 0) throw new Error(`AACで扱えないサンプルレートです：${sampleRate}`);
  const aot = 2; // AAC-LC
  return u8((aot << 3) | (idx >> 1), ((idx & 1) << 7) | ((channels & 15) << 3));
}

/** AACの設定情報から、サンプルレートとチャンネル数を読む（読めなければ null） */
export function parseAsc(asc) {
  if (!asc || asc.length < 2) return null;
  const aot = asc[0] >> 3;
  const idx = ((asc[0] & 7) << 1) | (asc[1] >> 7);
  const channels = (asc[1] >> 3) & 15;
  if (idx === 15 || !SAMPLE_RATES[idx]) return null;
  return { objectType: aot, sampleRate: SAMPLE_RATES[idx], channels };
}

/**
 * ADTS形式（1フレームごとに7〜9バイトの見出しが付いた形）を、フレームの配列に分ける。
 * 戻り値：{ frames: [Uint8Array]（見出しを除いた中身）, sampleRate, channels, asc }。ADTSでなければ null
 */
export function parseAdts(bytes) {
  const frames = [];
  let o = 0;
  let sampleRate = 0;
  let channels = 0;
  let profile = 1;
  while (o + 7 <= bytes.length) {
    if (bytes[o] !== 0xff || (bytes[o + 1] & 0xf6) !== 0xf0) break;
    const noCrc = bytes[o + 1] & 1;
    profile = bytes[o + 2] >> 6;
    const idx = (bytes[o + 2] >> 2) & 15;
    const ch = ((bytes[o + 2] & 1) << 2) | (bytes[o + 3] >> 6);
    const len = ((bytes[o + 3] & 3) << 11) | (bytes[o + 4] << 3) | (bytes[o + 5] >> 5);
    const head = noCrc ? 7 : 9;
    if (len < head || o + len > bytes.length || !SAMPLE_RATES[idx]) break;
    sampleRate = SAMPLE_RATES[idx];
    channels = ch;
    frames.push(bytes.subarray(o + head, o + len));
    o += len;
  }
  if (!frames.length || o !== bytes.length) return null;
  const idx = SAMPLE_RATES.indexOf(sampleRate);
  const asc = u8(((profile + 1) << 3) | (idx >> 1), ((idx & 1) << 7) | (channels << 3));
  return { frames, sampleRate, channels, asc };
}

const MATRIX = concat([u32(0x00010000), u32(0), u32(0), u32(0), u32(0x00010000), u32(0), u32(0), u32(0), u32(0x40000000)]);
const CHUNK_FRAMES = 32; // まとまり（チャンク）1つに入れるフレーム数

/**
 * AACのフレームからM4Aファイルを組み立てる。
 * frames：フレームの中身（Uint8Array）の配列。sampleRate・channels：音声の形式。asc：AACの設定情報（なければ自動で作る）。
 * bitrate：目安のビットレート（bps）。createdAt：録音した日時（ミリ秒。プレーヤーの「作成日時」に出る）
 * 戻り値：{ parts: [Uint8Array]（順に並べるとファイルになる）, bytes, durationSec, frames }
 */
export function buildM4a({ frames, sampleRate, channels = 1, asc = null, bitrate = 0, createdAt = 0 }) {
  if (!frames || !frames.length) throw new Error('音声のフレームがありません');
  const n = frames.length;
  const config = asc && asc.length >= 2 ? asc : ascFor(sampleRate, channels);
  const duration = n * AAC_FRAME;
  const created = createdAt ? Math.max(0, Math.floor(createdAt / 1000) + 2082844800) >>> 0 : 0; // 1904年からの秒
  let total = 0;
  let maxSize = 0;
  const stsz = new Uint8Array(12 + n * 4);
  const dvz = new DataView(stsz.buffer);
  dvz.setUint32(8, n);
  for (let i = 0; i < n; i++) {
    const len = frames[i].length;
    dvz.setUint32(12 + i * 4, len);
    total += len;
    if (len > maxSize) maxSize = len;
  }
  const avg = bitrate || Math.round((total * 8 * sampleRate) / duration);
  const nChunks = Math.ceil(n / CHUNK_FRAMES);
  const lastCount = n - (nChunks - 1) * CHUNK_FRAMES;
  const stscEntries = [];
  if (nChunks > 1) stscEntries.push([1, CHUNK_FRAMES, 1]);
  if (nChunks === 1 || lastCount !== CHUNK_FRAMES) stscEntries.push([nChunks, lastCount, 1]);

  const makeMoov = (dataStart) => {
    const stco = new Uint8Array(8 + nChunks * 4);
    const dvc = new DataView(stco.buffer);
    dvc.setUint32(4, nChunks);
    let off = dataStart;
    for (let c = 0, i = 0; c < nChunks; c++) {
      dvc.setUint32(8 + c * 4, off);
      for (let k = 0; k < CHUNK_FRAMES && i < n; k++, i++) off += frames[i].length;
    }
    const esds = fullBox('esds', 0, 0,
      descriptor(0x03, u16(1), u8(0),
        descriptor(0x04, u8(0x40, 0x15), u8(0, (maxSize >>> 8) & 255, maxSize & 255), u32(Math.max(avg, 1)), u32(Math.max(avg, 1)),
          descriptor(0x05, config)),
        descriptor(0x06, u8(2))));
    const mp4a = box('mp4a', zeros(6), u16(1), zeros(8), u16(channels), u16(16), u16(0), u16(0), u32((sampleRate << 16) >>> 0), esds);
    const stbl = box('stbl',
      fullBox('stsd', 0, 0, u32(1), mp4a),
      fullBox('stts', 0, 0, u32(1), u32(n), u32(AAC_FRAME)),
      fullBox('stsc', 0, 0, u32(stscEntries.length), ...stscEntries.flatMap((e) => e.map(u32))),
      box('stsz', stsz),
      box('stco', stco));
    const minf = box('minf',
      fullBox('smhd', 0, 0, u16(0), u16(0)),
      box('dinf', fullBox('dref', 0, 0, u32(1), fullBox('url ', 0, 1))),
      stbl);
    const mdia = box('mdia',
      fullBox('mdhd', 0, 0, u32(created), u32(created), u32(sampleRate), u32(duration), u16(0x55c4), u16(0)),
      fullBox('hdlr', 0, 0, u32(0), str('soun'), zeros(12), str('SoundHandler\0')),
      minf);
    const trak = box('trak',
      fullBox('tkhd', 0, 7, u32(created), u32(created), u32(1), u32(0), u32(duration), zeros(8), u16(0), u16(0), u16(0x0100), u16(0), MATRIX, u32(0), u32(0)),
      mdia);
    return box('moov',
      fullBox('mvhd', 0, 0, u32(created), u32(created), u32(sampleRate), u32(duration), u32(0x00010000), u16(0x0100), zeros(10), MATRIX, zeros(24), u32(2)),
      trak);
  };

  const ftyp = box('ftyp', str('M4A '), u32(0), str('M4A '), str('mp42'), str('isom'));
  const size0 = makeMoov(0).length;
  const dataStart = ftyp.length + size0 + 8;
  const moov = makeMoov(dataStart);
  if (dataStart + total > 0xffffffff) throw new Error('録音が長すぎて、1つのM4Aファイルにできません');
  // 音声の中身は、約1MBずつにまとめる（細かい断片を大量に並べない）
  const parts = [ftyp, moov, concat([u32(total + 8), str('mdat')])];
  const BLOCK = 1 << 20;
  let buf = new Uint8Array(Math.min(BLOCK, total));
  let used = 0;
  for (const f of frames) {
    if (used + f.length > buf.length) {
      if (used) parts.push(buf.subarray(0, used));
      buf = new Uint8Array(Math.max(BLOCK, f.length));
      used = 0;
    }
    buf.set(f, used);
    used += f.length;
  }
  if (used) parts.push(buf.subarray(0, used));
  return { parts, bytes: dataStart + total, durationSec: duration / sampleRate, frames: n };
}

// ---------------------------------------------------------------------------
// 読み取り（このアプリが作ったM4Aを、書き起こしのために音声に戻すときに使う）
// ---------------------------------------------------------------------------
const tag = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);
const rd32 = (b, o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
const rd16 = (b, o) => (b[o] << 8) | b[o + 1];

/** 箱（box）の並びを順にたどる。fn(type, 中身の開始, 中身の終了) */
function eachBox(b, start, end, fn) {
  let o = start;
  while (o + 8 <= end) {
    let size = rd32(b, o);
    let head = 8;
    if (size === 1) { size = rd32(b, o + 8) * 4294967296 + rd32(b, o + 12); head = 16; }
    if (size === 0) size = end - o;
    if (size < head || o + size > end) break;
    fn(tag(b, o + 4), o + head, o + size);
    o += size;
  }
}

/** 記述子の並びから、AACの設定情報（タグ5）を探す */
function findAsc(b, start, end) {
  let o = start;
  while (o + 2 <= end) {
    const t = b[o++];
    let len = 0;
    for (let i = 0; i < 4 && o < end; i++) { const c = b[o++]; len = (len << 7) | (c & 0x7f); if (!(c & 0x80)) break; }
    const stop = Math.min(end, o + len);
    if (t === 0x05) return b.slice(o, stop);
    if (t === 0x03) { const hit = findAsc(b, o + 3, stop); if (hit) return hit; }
    else if (t === 0x04) { const hit = findAsc(b, o + 13, stop); if (hit) return hit; }
    o = stop;
  }
  return null;
}

/**
 * M4Aファイル（Blob）の目次を読む。
 * 戻り値：{ sampleRate, channels, asc, frameSamples, count, sizes, offsets, durationSec }（offsets はファイル先頭からの位置）
 */
export async function readM4a(blob) {
  // 先頭から箱をたどって、目次（moov）を探す
  let pos = 0;
  let moov = null;
  while (pos + 8 <= blob.size) {
    const h = new Uint8Array(await blob.slice(pos, pos + 16).arrayBuffer());
    let size = rd32(h, 0);
    if (size === 1 && h.length >= 16) size = rd32(h, 8) * 4294967296 + rd32(h, 12);
    if (size === 0) size = blob.size - pos;
    if (size < 8) break;
    if (tag(h, 4) === 'moov') { moov = new Uint8Array(await blob.slice(pos, pos + size).arrayBuffer()); break; }
    pos += size;
  }
  if (!moov) throw new Error('M4Aの目次（moov）が見つかりません');
  const info = { sampleRate: 0, channels: 1, asc: null, frameSamples: AAC_FRAME };
  let sizes = null;
  let fixedSize = 0;
  let count = 0;
  let stsc = [];
  let chunkOffsets = [];
  const walk = (start, end) => eachBox(moov, start, end, (type, a, z) => {
    if (['trak', 'mdia', 'minf', 'stbl'].includes(type)) { walk(a, z); return; }
    if (type === 'mdhd') info.sampleRate = moov[a] === 1 ? rd32(moov, a + 20) : rd32(moov, a + 12);
    else if (type === 'stsd') {
      eachBox(moov, a + 8, z, (t2, a2, z2) => {
        if (t2 !== 'mp4a') return;
        info.channels = rd16(moov, a2 + 16) || 1;
        eachBox(moov, a2 + 28, z2, (t3, a3, z3) => { if (t3 === 'esds') info.asc = findAsc(moov, a3 + 4, z3); });
      });
    } else if (type === 'stts') { if (rd32(moov, a + 4) >= 1) info.frameSamples = rd32(moov, a + 12) || AAC_FRAME; }
    else if (type === 'stsc') { const n = rd32(moov, a + 4); stsc = []; for (let i = 0; i < n; i++) stsc.push([rd32(moov, a + 8 + i * 12), rd32(moov, a + 12 + i * 12)]); }
    else if (type === 'stsz') {
      fixedSize = rd32(moov, a + 4);
      count = rd32(moov, a + 8);
      if (!fixedSize) { sizes = new Uint32Array(count); for (let i = 0; i < count; i++) sizes[i] = rd32(moov, a + 12 + i * 4); }
    } else if (type === 'stco') { const n = rd32(moov, a + 4); chunkOffsets = []; for (let i = 0; i < n; i++) chunkOffsets.push(rd32(moov, a + 8 + i * 4)); }
    else if (type === 'co64') { const n = rd32(moov, a + 4); chunkOffsets = []; for (let i = 0; i < n; i++) chunkOffsets.push(rd32(moov, a + 8 + i * 8) * 4294967296 + rd32(moov, a + 12 + i * 8)); }
  });
  eachBox(moov, 0, moov.length, (type, a, z) => { if (type === 'moov') walk(a, z); });
  if (!sizes) { sizes = new Uint32Array(count); sizes.fill(fixedSize); }
  const fromAsc = parseAsc(info.asc);
  if (fromAsc) { if (!info.sampleRate) info.sampleRate = fromAsc.sampleRate; }
  if (!info.asc || !info.sampleRate || !count || !chunkOffsets.length || !stsc.length) throw new Error('M4Aの目次を読み取れません（AACの音声ではない可能性があります）');
  // まとまり（チャンク）ごとの位置から、フレーム1つずつの位置を求める
  const offsets = new Float64Array(count);
  let i = 0;
  for (let c = 0; c < chunkOffsets.length && i < count; c++) {
    let per = stsc[0][1];
    for (const [first, n] of stsc) { if (c + 1 >= first) per = n; }
    let off = chunkOffsets[c];
    for (let k = 0; k < per && i < count; k++, i++) { offsets[i] = off; off += sizes[i]; }
  }
  if (i !== count) throw new Error('M4Aの目次が途中で終わっています');
  return { ...info, count, sizes, offsets, durationSec: (count * info.frameSamples) / info.sampleRate };
}
