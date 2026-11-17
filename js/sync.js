// 2台の同期（REQ-4 / B-7）：中継サーバーの代わりにOneDriveを使う。
// 録音する端末が _live.json と会議の状態を書き込み、もう一方の端末が数秒ごとに読みに行く。
// 各端末は自分のファイルだけに書き込むので、書き込みがぶつからない。
import * as od from './onedrive.js';
import { getSettings } from './settings.js';

export const rootPath = () => getSettings().onedrive.rootPath || '会議録音アーカイブ';
export const livePath = () => od.joinPath(rootPath(), '_live.json');
export const metaPath = (folderPath, name) => od.joinPath(folderPath, '_meta', name);
export const standbyPath = () => od.joinPath(rootPath(), '_standby.json'); // 待機中の端末（ほかの端末からの録音開始を待つ）
export const startPath = () => od.joinPath(rootPath(), '_start.json');     // 待機中の端末への、録音開始の指示

/** 変更があったときだけ JSON を読む */
export async function readJsonIfChanged(path, lastETag) {
  const item = await od.getItem(path);
  if (!item) return { missing: true };
  if (lastETag && item.eTag === lastETag) return { unchanged: true, eTag: item.eTag };
  const url = item['@microsoft.graph.downloadUrl'];
  if (!url) return { missing: true };
  const r = await fetch(url);
  if (!r.ok) throw new Error(`読み込み失敗 ${r.status}`);
  const data = await r.json().catch(() => null);
  return { data, eTag: item.eTag };
}

// ---------------------------------------------------------------------------
// 録音する端末側：状況の発信と、遠隔停止の受信
// ---------------------------------------------------------------------------
export class LivePublisher {
  constructor({ session, getElapsed, onRemoteStop, onRemoteProcess, onError }) {
    this.session = session;
    this.getElapsed = getElapsed;
    this.onRemoteStop = onRemoteStop;
    this.onRemoteProcess = onRemoteProcess || (() => {});
    this.onError = onError || (() => {});
    this.phase = 'recording'; // recording → stopping → recorded（保存するかの確認待ち）→ ended
    this.controlETag = null;
  }

  liveRecord(status) {
    const s = this.session;
    const st = getSettings();
    return {
      sessionId: s.id, folderPath: s.folderPath, folderName: s.folderName, title: s.title,
      projectId: s.projectId, startedAt: s.startedAt, status,
      deviceId: st.deviceId, deviceName: st.deviceName,
      elapsed: this.getElapsed ? this.getElapsed() : 0, heartbeat: Date.now(),
    };
  }

  async start() {
    await this.write('recording');
    this.hb = setInterval(() => this.write('recording'), 15000);
    // 停止指示の確認は8秒ごと（OneDriveへの問い合わせ回数を抑える）
    const poll = Math.max(8, Number(getSettings().sync.pollSec) || 5) * 1000;
    this.ctl = setInterval(() => this.checkControl(), poll);
  }

  async write(status) {
    try { await od.uploadJson(livePath(), this.liveRecord(status)); } catch (e) { this.onError(e); }
  }

  async checkControl() {
    if (this.phase === 'ended' || this.phase === 'stopping' || !this.session.folderPath) return;
    try {
      const res = await readJsonIfChanged(metaPath(this.session.folderPath, 'control.json'), this.controlETag);
      if (res.eTag) this.controlETag = res.eTag;
      const c = res.data;
      if (!c || c.sessionId !== this.session.id) return;
      if (c.command === 'stop' && this.phase === 'recording') {
        this.phase = 'stopping';
        this.onRemoteStop(c);
      } else if (c.command === 'process' && this.phase === 'recorded') {
        this.phase = 'stopping';
        this.onRemoteProcess(c);
      }
    } catch (e) { this.onError(e); }
  }

  /** 録音を終え、保存するかの確認を待つ間：状態を「録音済み」と発信し、もう一方の端末からの保存の指示を待つ */
  async hold() {
    this.phase = 'recorded';
    clearInterval(this.hb);
    this.hb = setInterval(() => this.write('recorded'), 15000);
    await this.write('recorded');
  }

  /** 何も書き込まずに終える（次の会議の録音を始めるとき。前の会議の状態は上書きされる） */
  drop() {
    this.phase = 'ended';
    clearInterval(this.hb);
    clearInterval(this.ctl);
  }

  async stop(status = 'stopped') {
    this.phase = 'ended';
    clearInterval(this.hb);
    clearInterval(this.ctl);
    await this.write(status);
  }
}

// ---------------------------------------------------------------------------
// もう一方の端末側：録音中の会議を見つける
// ---------------------------------------------------------------------------
export class LiveWatcher {
  constructor({ onChange, onStandby, onError }) {
    this.onChange = onChange;
    this.onStandby = onStandby || (() => {});
    this.onError = onError || (() => {});
    this.eTag = null;
    this.current = null;
    this.timer = null;
    this.standbyETag = null;
    this.standby = null;
  }

  start() {
    if (this.timer) return;
    const tick = async () => {
      if (document.visibilityState === 'visible') { await this.check(); await this.checkStandby(); }
      const poll = Math.max(3, Number(getSettings().sync.pollSec) || 5) * 1000;
      this.timer = setTimeout(tick, poll);
    };
    this.timer = setTimeout(tick, 100);
  }

  stop() { clearTimeout(this.timer); this.timer = null; }

  /** ほかの端末が「待機中」（こちらから録音を開始できる状態）かを調べる */
  async checkStandby() {
    try {
      const res = await readJsonIfChanged(standbyPath(), this.standbyETag);
      if (res.eTag) this.standbyETag = res.eTag;
      let sb = res.unchanged ? this.standbyRaw : (res.data || null);
      this.standbyRaw = sb;
      const mine = sb && sb.deviceId === getSettings().deviceId;
      // 待機中の端末は20秒ごとに発信する。70秒以上途絶えたら、待機していないとみなす
      const fresh = sb && sb.status === 'standby' && Date.now() - sb.heartbeat < 70_000;
      sb = sb && !mine && fresh ? sb : null;
      const key = (x) => (x ? `${x.deviceId}|${x.deviceName}|${x.lastScene}` : '');
      const changed = key(sb) !== key(this.standby);
      this.standby = sb;
      if (changed) this.onStandby(sb);
    } catch (e) { this.onError(e); }
  }

  async check() {
    try {
      const res = await readJsonIfChanged(livePath(), this.eTag);
      if (res.eTag) this.eTag = res.eTag;
      if (res.unchanged) {
        // 状態に変化がなくても、発信が途絶えた（端末の電池切れ等）かは確認する
        if (this.current && this.current.status === 'recording' && Date.now() - this.current.heartbeat > 90_000) {
          this.current = { ...this.current, status: 'lost' };
          this.onChange(this.current);
        }
        return;
      }
      const live = res.data || null;
      const mine = live && live.deviceId === getSettings().deviceId;
      // 録音中は90秒以内の発信があるもの、録音終了後は6時間以内のものだけを対象にする
      const fresh = live && (live.status === 'recording' ? Date.now() - live.heartbeat < 90_000 : Date.now() - live.heartbeat < 6 * 3600_000);
      const next = live && !mine && fresh ? live : null;
      // 定期的な発信（経過時間・時刻）だけの変化では画面を作り直さない
      const key = (x) => (x ? `${x.sessionId}|${x.status}|${x.folderPath}|${x.title}|${x.deviceName}|${x.startedAt}` : '');
      const changed = key(next) !== key(this.current);
      this.current = next;
      if (changed) this.onChange(next);
    } catch (e) { this.onError(e); }
  }
}

// ---------------------------------------------------------------------------
// 待機：この端末を、ほかの端末から録音開始できる状態にする
// ---------------------------------------------------------------------------
export class StandbyAgent {
  constructor({ onStart, onError, info }) {
    this.onStart = onStart;
    this.onError = onError || (() => {});
    this.info = info || (() => ({}));
    this.active = false;
    this.startETag = null;
  }

  record(status) {
    const st = getSettings();
    return { deviceId: st.deviceId, deviceName: st.deviceName, status, heartbeat: Date.now(), ...this.info() };
  }

  async start() {
    this.active = true;
    // 待機を始める前からある指示は無視する（その時点の版を覚えておき、変わったときだけ読む）
    try { const item = await od.getItem(startPath()); this.startETag = item ? item.eTag : null; } catch (e) { this.onError(e); }
    await od.uploadJson(standbyPath(), this.record('standby'));
    this.hb = setInterval(() => { od.uploadJson(standbyPath(), this.record('standby')).catch((e) => this.onError(e)); }, 20000);
    this.poll = setInterval(() => this.check(), 4000);
  }

  async check() {
    if (!this.active || this.busy) return;
    this.busy = true;
    try {
      const res = await readJsonIfChanged(startPath(), this.startETag);
      if (res.eTag) this.startETag = res.eTag;
      const c = res.data;
      if (c && c.command === 'start' && c.target === getSettings().deviceId && this.active) {
        await this.stop('starting');
        this.onStart(c);
      }
    } catch (e) { this.onError(e); } finally { this.busy = false; }
  }

  async stop(status = 'off') {
    if (!this.active) return;
    this.active = false;
    clearInterval(this.hb);
    clearInterval(this.poll);
    try { await od.uploadJson(standbyPath(), this.record(status)); } catch (e) { this.onError(e); }
  }
}

/** 待機中の端末に、録音の開始を指示する */
export async function requestRemoteStart(standby, params) {
  const st = getSettings();
  await od.uploadJson(startPath(), {
    command: 'start', target: standby.deviceId, params, at: Date.now(), by: st.deviceName, byDevice: st.deviceId,
    nonce: Math.random().toString(36).slice(2),
  });
}

/** 録音状況のファイルに、1回だけ状態を書く（保存完了の通知など） */
export async function writeLiveOnce(session, status) {
  const st = getSettings();
  await od.uploadJson(livePath(), {
    sessionId: session.id, folderPath: session.folderPath, folderName: session.folderName, title: session.title,
    projectId: session.projectId, startedAt: session.startedAt, status,
    deviceId: session.deviceId || st.deviceId, deviceName: session.deviceName || st.deviceName, elapsed: session.durationSec || 0, heartbeat: Date.now(),
  });
}

/** 録音している端末に指示する：'stop'＝録音の終了、'process'＝①（保存・書き起こし）の開始 */
export async function requestRemote(live, command) {
  const st = getSettings();
  await od.uploadJson(metaPath(live.folderPath, 'control.json'), {
    command, sessionId: live.sessionId, at: Date.now(), by: st.deviceName, byDevice: st.deviceId,
  });
}
export const requestRemoteStop = (live) => requestRemote(live, 'stop');

/** 会議フォルダ内の、全端末のメモを読む */
export async function readAllMemos(folderPath) {
  const items = await od.listChildren(od.joinPath(folderPath, '_meta'));
  const files = items.filter((i) => /^memos_.+\.json$/.test(i.name));
  const lists = await Promise.all(files.map((f) => od.readJson(od.joinPath(folderPath, '_meta', f.name)).catch(() => null)));
  return lists.filter(Boolean).map((x) => x.memos || []);
}

export async function writeMyMemos(folderPath, memos) {
  const st = getSettings();
  await od.uploadJson(metaPath(folderPath, `memos_${st.deviceId}.json`), {
    deviceId: st.deviceId, deviceName: st.deviceName, updatedAt: Date.now(),
    memos: memos.filter((m) => m.deviceId === st.deviceId),
  });
}
