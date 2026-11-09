// 録音用 AudioWorklet：入力をモノラル化し、16kHz に変換して本体へ送る。
// 端末のマイクは 48kHz / 44.1kHz などさまざまなので、ここで平均化しながら間引く。

const TARGET_RATE = 16000;
const BATCH = 4000; // 0.25秒分ずつ送る

class PcmCapture extends AudioWorkletProcessor {
  constructor() {
    super();
    this.ratio = sampleRate / TARGET_RATE; // sampleRate は AudioWorkletGlobalScope の値
    this.phase = 0;
    this.sum = 0;
    this.count = 0;
    this.buf = new Float32Array(BATCH);
    this.len = 0;
    this.active = true;
    this.port.onmessage = (e) => {
      if (e.data === 'flush') this.flush();
      if (e.data === 'stop') { this.flush(); this.active = false; }
    };
  }

  flush() {
    if (this.len > 0) {
      const out = this.buf.slice(0, this.len);
      this.port.postMessage(out, [out.buffer]);
      this.len = 0;
    }
  }

  process(inputs) {
    if (!this.active) return false;
    const input = inputs[0];
    if (!input || input.length === 0 || !input[0]) return true;
    const chs = input.length;
    const frames = input[0].length;
    for (let i = 0; i < frames; i++) {
      let x = 0;
      for (let c = 0; c < chs; c++) x += input[c][i];
      x /= chs;
      this.sum += x;
      this.count++;
      this.phase += 1;
      while (this.phase >= this.ratio) { // 入力が16kHz未満の場合は同じ値を繰り返して補う
        this.phase -= this.ratio;
        const v = this.count ? this.sum / this.count : (this.last || 0);
        this.last = v;
        this.buf[this.len++] = v;
        this.sum = 0;
        this.count = 0;
        if (this.len === BATCH) this.flush();
      }
    }
    return true;
  }
}

registerProcessor('pcm-capture', PcmCapture);
