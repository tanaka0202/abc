// WAV（16bit PCM・モノラル）の組み立て。音声データはメモリに全部載せず Blob のまま扱う。

export const SAMPLE_RATE = 16000;
export const BYTES_PER_SAMPLE = 2;

/** WAVヘッダ（44バイト） */
export function wavHeader(numSamples, sampleRate = SAMPLE_RATE, channels = 1) {
  const dataBytes = numSamples * BYTES_PER_SAMPLE * channels;
  const buf = new ArrayBuffer(44);
  const v = new DataView(buf);
  const str = (off, s) => { for (let i = 0; i < s.length; i++) v.setUint8(off + i, s.charCodeAt(i)); };
  str(0, 'RIFF');
  v.setUint32(4, 36 + dataBytes, true);
  str(8, 'WAVE');
  str(12, 'fmt ');
  v.setUint32(16, 16, true);              // fmtチャンクの長さ
  v.setUint16(20, 1, true);               // PCM
  v.setUint16(22, channels, true);
  v.setUint32(24, sampleRate, true);
  v.setUint32(28, sampleRate * channels * BYTES_PER_SAMPLE, true);
  v.setUint16(32, channels * BYTES_PER_SAMPLE, true);
  v.setUint16(34, 16, true);
  str(36, 'data');
  v.setUint32(40, dataBytes, true);
  return buf;
}

/** Float32（-1〜1） → Int16 */
export function floatToInt16(float32) {
  const out = new Int16Array(float32.length);
  for (let i = 0; i < float32.length; i++) {
    const x = Math.max(-1, Math.min(1, float32[i]));
    out[i] = x < 0 ? x * 0x8000 : x * 0x7fff;
  }
  return out;
}

/** 録音の断片（Int16のArrayBuffer配列）から、ヘッダなしのPCM Blob */
export function pcmBlobFromChunks(chunks) {
  return new Blob(chunks, { type: 'application/octet-stream' });
}

/** PCM Blob → WAV Blob（全体） */
export function wavFromPcmBlob(pcmBlob, sampleRate = SAMPLE_RATE) {
  const n = Math.floor(pcmBlob.size / BYTES_PER_SAMPLE);
  return new Blob([wavHeader(n, sampleRate), pcmBlob], { type: 'audio/wav' });
}

/** PCM Blob の一部（秒で指定）→ WAV Blob */
export function wavSlice(pcmBlob, startSec, endSec, sampleRate = SAMPLE_RATE) {
  const total = Math.floor(pcmBlob.size / BYTES_PER_SAMPLE);
  const a = Math.max(0, Math.min(total, Math.floor(startSec * sampleRate)));
  const b = Math.max(a, Math.min(total, Math.floor(endSec * sampleRate)));
  const part = pcmBlob.slice(a * BYTES_PER_SAMPLE, b * BYTES_PER_SAMPLE);
  return new Blob([wavHeader(b - a, sampleRate), part], { type: 'audio/wav' });
}

/** WAVファイル（Blob）からPCM部分とサンプルレートを取り出す（取り込んだ音声の再処理用） */
export async function parseWav(blob) {
  const head = new DataView(await blob.slice(0, Math.min(blob.size, 4096)).arrayBuffer());
  const tag = (o) => String.fromCharCode(head.getUint8(o), head.getUint8(o + 1), head.getUint8(o + 2), head.getUint8(o + 3));
  if (tag(0) !== 'RIFF' || tag(8) !== 'WAVE') throw new Error('WAV形式ではありません');
  let off = 12;
  let sampleRate = SAMPLE_RATE;
  let channels = 1;
  let bits = 16;
  while (off + 8 <= head.byteLength) {
    const id = tag(off);
    const size = head.getUint32(off + 4, true);
    if (id === 'fmt ') {
      channels = head.getUint16(off + 10, true);
      sampleRate = head.getUint32(off + 12, true);
      bits = head.getUint16(off + 22, true);
    } else if (id === 'data') {
      if (bits !== 16 || channels !== 1) throw new Error('16bit・モノラルのWAVのみ対応しています');
      const start = off + 8;
      return { pcm: blob.slice(start, Math.min(blob.size, start + size)), sampleRate };
    }
    off += 8 + size + (size % 2);
  }
  throw new Error('WAVのデータ部分が見つかりません');
}

/** 1秒ごとの音量（dBFS）を計算する（Int16配列） */
export function rmsDb(int16) {
  if (!int16.length) return -100;
  let sum = 0;
  for (let i = 0; i < int16.length; i++) { const x = int16[i] / 32768; sum += x * x; }
  const rms = Math.sqrt(sum / int16.length);
  return rms > 0 ? Math.max(-100, 20 * Math.log10(rms)) : -100;
}
