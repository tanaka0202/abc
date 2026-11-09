// 録音：マイク（とPCではタブ音声）を16kHzモノラルにして、5秒ごとに端末内へ退避する。
import { addChunk, putSession } from './db.js';
import { floatToInt16, SAMPLE_RATE } from './wav.js';
import { AppError } from './util.js';

const FLUSH_SAMPLES = SAMPLE_RATE * 5;      // 5秒ごとに保存
const GAP_THRESHOLD = SAMPLE_RATE * 2;      // 2秒以上の欠落を途切れとみなす
const GAP_SLACK = SAMPLE_RATE * 0.3;        // 処理の遅れの分は補わない

export class Recorder {
  constructor({ session, settings, onLevel, onTick, onWarning, onStarted }) {
    this.session = session;
    this.settings = settings;
    this.onLevel = onLevel || (() => {});
    this.onTick = onTick || (() => {});
    this.onWarning = onWarning || (() => {});
    this.onStarted = onStarted || (() => {});
    this.streams = [];
    this.pending = [];
    this.pendingSamples = 0;
    this.totalSamples = 0;
    this.seq = 0;
    this.levels = [];
    this.gaps = [];
    this.secSum = 0;
    this.secCount = 0;
    this.t0 = null;
    this.writeChain = Promise.resolve();
    this.writeErrors = 0;
    this.lowWarned = false;
    this.running = false;
    this.onVisibility = this.onVisibility.bind(this);
  }

  async start(source = 'mic') {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      throw new AppError('このブラウザでは録音できません（HTTPSで開いているか確認してください）');
    }
    const rs = this.settings.recording;
    // 画面共有の許可は操作の直後でないと求められないため、タブ音声を先に取得する
    if (source === 'tab') {
      try {
        const disp = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: true, systemAudio: 'include' });
        disp.getVideoTracks().forEach((t) => { t.enabled = false; });
        if (!disp.getAudioTracks().length) {
          this.onWarning('noTabAudio', '画面共有で「音声を共有」が選ばれていないため、マイクの音だけを録音します。');
        }
        this.streams.push(disp);
      } catch {
        this.onWarning('noTabAudio', 'タブの音声を取得できなかったため、マイクの音だけを録音します。');
      }
    }

    let mic;
    try {
      mic = await navigator.mediaDevices.getUserMedia({
        audio: {
          echoCancellation: !!rs.echoCancellation,
          noiseSuppression: !!rs.noiseSuppression,
          autoGainControl: !!rs.autoGainControl,
          channelCount: 1,
        },
      });
    } catch (e) {
      for (const st of this.streams) st.getTracks().forEach((t) => t.stop()); // 取得済みの画面共有を止める
      this.streams = [];
      throw new AppError('マイクを使えませんでした。ブラウザのマイクの許可を確認してください。', { cause: e });
    }
    this.streams.push(mic);

    const Ctx = window.AudioContext || window.webkitAudioContext;
    this.ctx = new Ctx({ latencyHint: 'playback' });
    await this.ctx.audioWorklet.addModule(new URL('./audio-worklet.js', import.meta.url));
    this.node = new AudioWorkletNode(this.ctx, 'pcm-capture', { numberOfInputs: 1, numberOfOutputs: 1, outputChannelCount: [1] });
    this.node.port.onmessage = (e) => this.onAudio(e.data);

    let entry = this.node;
    if (rs.highpass) {
      const hp = this.ctx.createBiquadFilter();
      hp.type = 'highpass';
      hp.frequency.value = 90; // 空調・机の振動などの低い雑音を除く
      hp.connect(this.node);
      entry = hp;
    }
    for (const s of this.streams) {
      if (!s.getAudioTracks().length) continue;
      this.ctx.createMediaStreamSource(s).connect(entry);
      s.getAudioTracks().forEach((t) => t.addEventListener('ended', () => {
        if (this.running) this.onWarning('ended', '音声の入力が止まりました（マイクが他のアプリに使われた可能性があります）。');
      }));
    }
    // AudioWorklet を確実に動かすため、無音で出力につなぐ
    const mute = this.ctx.createGain();
    mute.gain.value = 0;
    this.node.connect(mute).connect(this.ctx.destination);

    this.ctx.onstatechange = () => {
      if (this.running && this.ctx.state !== 'running') {
        this.ctx.resume().catch(() => {});
        this.onWarning('suspended', '録音が一時停止された可能性があります。画面をつけたままにしてください。');
      }
    };
    if (this.ctx.state === 'suspended') await this.ctx.resume();

    this.running = true;
    this.timer = setInterval(() => this.tick(), 1000);
    document.addEventListener('visibilitychange', this.onVisibility);
    await this.acquireWakeLock();
  }

  onAudio(float32) {
    if (!this.running || !float32 || !float32.length) return;
    const n = float32.length;
    const now = Date.now();
    if (this.t0 === null) {
      this.t0 = now - (n / SAMPLE_RATE) * 1000;
      this.onStarted(this.t0);
    } else {
      // 実際の経過時間より録音が短い＝途切れた。無音で補い、メモとの時刻のずれを防ぐ
      const expected = ((now - this.t0) / 1000) * SAMPLE_RATE;
      const lag = expected - (this.totalSamples + n);
      if (lag > GAP_THRESHOLD) {
        const pad = Math.floor(lag - GAP_SLACK);
        this.pushSamples(new Int16Array(pad));
        this.gaps.push({ at: Math.round(this.totalSamples / SAMPLE_RATE), sec: Math.round(pad / SAMPLE_RATE) });
        this.onWarning('gap', `録音が約${Math.round(pad / SAMPLE_RATE)}秒途切れました（無音で補いました）。`);
      }
    }
    this.pushSamples(floatToInt16(float32));
  }

  pushSamples(int16) {
    this.pending.push(int16);
    this.pendingSamples += int16.length;
    this.totalSamples += int16.length;
    // 1秒ごとの音量
    for (let i = 0; i < int16.length; i++) {
      const x = int16[i] / 32768;
      this.secSum += x * x;
      this.secCount++;
      if (this.secCount >= SAMPLE_RATE) {
        const rms = Math.sqrt(this.secSum / this.secCount);
        const db = rms > 0 ? Math.max(-100, 20 * Math.log10(rms)) : -100;
        this.levels.push(Math.round(db * 10) / 10);
        this.onLevel(db);
        this.secSum = 0;
        this.secCount = 0;
      }
    }
    if (this.pendingSamples >= FLUSH_SAMPLES) this.flush();
  }

  flush() {
    if (!this.pendingSamples) return this.writeChain;
    const merged = new Int16Array(this.pendingSamples);
    let off = 0;
    for (const p of this.pending) { merged.set(p, off); off += p.length; }
    this.pending = [];
    this.pendingSamples = 0;
    const seq = this.seq++;
    const sessionId = this.session.id;
    this.writeChain = this.writeChain
      .then(() => addChunk(sessionId, seq, merged.buffer))
      .catch((e) => {
        this.writeErrors++;
        console.error(e);
        this.onWarning('storage', '端末内への保存に失敗しました。空き容量を確認してください。');
      });
    if (seq % 6 === 5) this.saveProgress(); // 30秒ごとに会議情報も更新（復元用）
    return this.writeChain;
  }

  saveProgress() {
    this.session.durationSec = this.totalSamples / SAMPLE_RATE;
    this.session.levels = this.levels.slice();
    this.session.gaps = this.gaps.slice();
    this.session.chunkCount = this.seq;
    putSession(this.session).catch(() => {});
  }

  tick() {
    const elapsed = this.totalSamples / SAMPLE_RATE;
    this.onTick(elapsed);
    if (!this.lowWarned && this.levels.length >= 15) {
      const recent = this.levels.slice(-15).slice().sort((a, b) => a - b);
      const median = recent[Math.floor(recent.length / 2)];
      if (median < -50) {
        this.lowWarned = true;
        this.onWarning('low', '音が小さいようです。PCのスピーカー音量を上げるか、スマホを話者に近づけてください。');
      }
    }
  }

  get elapsed() { return this.totalSamples / SAMPLE_RATE; }

  async acquireWakeLock() {
    try {
      if ('wakeLock' in navigator) {
        this.wake = await navigator.wakeLock.request('screen');
      }
    } catch { /* 省電力モードなどで取得できない場合がある */ }
  }

  onVisibility() {
    if (!this.running) return;
    if (document.visibilityState === 'visible') {
      this.acquireWakeLock();
      if (this.ctx && this.ctx.state !== 'running') this.ctx.resume().catch(() => {});
    } else {
      this.onWarning('hidden', '録音中は画面をつけたまま、このアプリを表示しておいてください。');
    }
  }

  async stop() {
    if (!this.running) return this.result();
    try { this.node.port.postMessage('stop'); } catch { /* 無視 */ }
    await new Promise((r) => setTimeout(r, 400)); // 最後のデータを受け取る
    this.running = false;
    clearInterval(this.timer);
    document.removeEventListener('visibilitychange', this.onVisibility);
    this.flush();
    try { this.node.disconnect(); } catch { /* 無視 */ }
    try { await this.ctx.close(); } catch { /* 無視 */ }
    for (const s of this.streams) s.getTracks().forEach((t) => t.stop());
    try { if (this.wake) await this.wake.release(); } catch { /* 無視 */ }
    await this.writeChain;
    this.saveProgress();
    return this.result();
  }

  result() {
    return {
      startedAt: this.t0,
      durationSec: this.totalSamples / SAMPLE_RATE,
      levels: this.levels.slice(),
      gaps: this.gaps.slice(),
      chunkCount: this.seq,
      writeErrors: this.writeErrors,
    };
  }
}
