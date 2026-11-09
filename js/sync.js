// 2台の同期（REQ-4 / B-7）：中継サーバーの代わりにOneDriveを使う。
// 録音する端末が _live.json と会議の状態を書き込み、もう一方の端末が数秒ごとに読みに行く。
// 各端末は自分のファイルだけに書き込むので、書き込みがぶつからない。
import * as od from './onedrive.js';
import { getSettings } from './settings.js';

export const rootPath = () => getSettings().onedrive.rootPath || '会議録音アーカイブ';
export const livePath = () => od.joinPath(rootPath(), '_live.json');
export const metaPath = (folderPath, name) => od.joinPath(folderPath, '_meta', name);

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
  constructor({ session, getElapsed, onRemoteStop, onError }) {
    this.session = session;
    this.getElapsed = getElapsed;
    this.onRemoteStop = onRemoteStop;
    this.onError = onError || (() => {});
    this.stopped = false;
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
    if (this.stopped || !this.session.folderPath) return;
    try {
      const res = await readJsonIfChanged(metaPath(this.session.folderPath, 'control.json'), this.controlETag);
      if (res.eTag) this.controlETag = res.eTag;
      const c = res.data;
      if (c && c.command === 'stop' && c.sessionId === this.session.id && !this.stopped) {
        this.stopped = true;
        this.onRemoteStop(c);
      }
    } catch (e) { this.onError(e); }
  }

  async stop(status = 'stopped') {
    this.stopped = true;
    clearInterval(this.hb);
    clearInterval(this.ctl);
    await this.write(status);
  }
}

// ---------------------------------------------------------------------------
// もう一方の端末側：録音中の会議を見つける
// ---------------------------------------------------------------------------
export class LiveWatcher {
  constructor({ onChange, onError }) {
    this.onChange = onChange;
    this.onError = onError || (() => {});
    this.eTag = null;
    this.current = null;
    this.timer = null;
  }

  start() {
    if (this.timer) return;
    const tick = async () => {
      if (document.visibilityState === 'visible') await this.check();
      const poll = Math.max(3, Number(getSettings().sync.pollSec) || 5) * 1000;
      this.timer = setTimeout(tick, poll);
    };
    this.timer = setTimeout(tick, 100);
  }

  stop() { clearTimeout(this.timer); this.timer = null; }

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

/** 録音している端末に停止を指示する */
export async function requestRemoteStop(live) {
  const st = getSettings();
  await od.uploadJson(metaPath(live.folderPath, 'control.json'), {
    command: 'stop', sessionId: live.sessionId, at: Date.now(), by: st.deviceName, byDevice: st.deviceId,
  });
}

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
