// 録音中・処理中に画面を黒くする（案A）。2回タップで操作を出し、5秒操作しなければ黒に戻る。
// 録音が止まったときだけ、画面を明るくして振動で知らせる。
import { el, clear, fmtTime } from './util.js';
import { holdButton } from './ui.js';

const DOUBLE_TAP_MS = 400;
const DOUBLE_TAP_PX = 48;
const PANEL_IDLE_MS = 5000;
const HINT_MS = 3000;

const ICONS = {
  star: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"/></svg>',
  memo: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M4 20h4l10-10-4-4L4 16z"/><path d="M13 7l4 4"/></svg>',
  moon: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/></svg>',
  back: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M4 12h16"/><path d="M10 6l-6 6 6 6"/></svg>',
  warn: '<svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3l10 18H2z"/><path d="M12 10v5"/><path d="M12 18v.5"/></svg>',
};

function iconButton(icon, label, onClick, extra = {}) {
  const b = el('button', { class: 'bo-btn', type: 'button', onclick: onClick, ...extra });
  const i = document.createElement('span');
  i.className = 'bo-ico';
  i.innerHTML = ICONS[icon]; // 固定のSVG（利用者の入力は含まない）
  b.append(i, el('span', {}, label));
  return b;
}

export class Blackout {
  /**
   * getStatus()：{ mode: 'rec'|'proc', elapsed, saved, progressText, progressRatio, notes: [] }
   * onPin / onMemo / onStop / onExit：各ボタンの処理
   */
  constructor({ getStatus, onPin, onMemo, onStop, onExit, holdToStop = true }) {
    this.getStatus = getStatus;
    this.onPin = onPin;
    this.onMemo = onMemo;
    this.onStop = onStop;
    this.onExit = onExit;
    this.holdToStop = holdToStop;
    this.root = null;
    this.alertBox = null;
    this.lastTap = null;
    this.idleTimer = null;
    this.tickTimer = null;
    this.flash = '';
  }

  get active() { return !!this.root; }

  /** 黒い画面にする。利用者の操作の直後なら、全画面にして上下の帯も消す */
  show({ hint = true } = {}) {
    if (this.root) return;
    this.root = el('div', { class: 'blackout', tabindex: '0', 'aria-label': '暗い画面。2回タップ（またはEnter）で操作を表示します' });
    this.dot = el('div', { class: 'bo-dot', 'aria-hidden': 'true' });
    this.hint = hint ? el('p', { class: 'bo-hint' }, '2回タップすると操作が表示されます') : null;
    this.panel = el('div', { class: 'bo-panel', hidden: true });
    this.root.append(this.dot, ...(this.hint ? [this.hint] : []), this.panel);
    this.root.addEventListener('pointerup', (e) => this.onTap(e));
    this.root.addEventListener('keydown', (e) => {
      if (e.target !== this.root) return;
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.showPanel(); }
      if (e.key === 'Escape') { e.preventDefault(); this.exit(); }
    });
    this.root.addEventListener('contextmenu', (e) => e.preventDefault());
    document.body.appendChild(this.root);
    document.body.classList.add('bo-on');
    this.root.focus({ preventScroll: true });
    if (this.hint) setTimeout(() => { if (this.hint) this.hint.classList.add('bo-fade'); }, HINT_MS);
    try {
      const fs = document.documentElement.requestFullscreen;
      if (fs && !document.fullscreenElement) document.documentElement.requestFullscreen({ navigationUI: 'hide' }).catch(() => {});
    } catch { /* 全画面にできない環境では黒い画面だけにする */ }
  }

  /** 黒い画面をやめて、通常の画面に戻す */
  hide() {
    this.hidePanel();
    if (this.root) { this.root.remove(); this.root = null; }
    document.body.classList.remove('bo-on');
    try { if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); } catch { /* 無視 */ }
  }

  exit() {
    this.hide();
    if (this.onExit) this.onExit();
  }

  onTap(e) {
    if (!this.root || (this.panel && !this.panel.hidden && this.panel.contains(e.target))) return;
    const now = performance.now();
    const p = this.lastTap;
    if (p && now - p.t < DOUBLE_TAP_MS && Math.hypot(e.clientX - p.x, e.clientY - p.y) < DOUBLE_TAP_PX) {
      this.lastTap = null;
      if (this.panel.hidden) this.showPanel(); else this.hidePanel();
      return;
    }
    this.lastTap = { t: now, x: e.clientX, y: e.clientY };
  }

  showPanel() {
    if (!this.root) return;
    this.panel.hidden = false;
    this.root.classList.add('bo-panel-on');
    if (this.hint) { this.hint.remove(); this.hint = null; }
    this.renderPanel();
    clearInterval(this.tickTimer);
    this.tickTimer = setInterval(() => this.renderPanelStatus(), 1000);
    this.bumpIdle();
  }

  hidePanel() {
    clearTimeout(this.idleTimer);
    clearInterval(this.tickTimer);
    if (this.panel) this.panel.hidden = true;
    if (this.root) { this.root.classList.remove('bo-panel-on'); this.root.focus({ preventScroll: true }); }
  }

  /** 操作がなければ5秒で黒に戻す（長押しの途中は戻さない） */
  bumpIdle() {
    clearTimeout(this.idleTimer);
    this.idleTimer = setTimeout(() => {
      if (this.panel && this.panel.querySelector('.holding')) { this.bumpIdle(); return; }
      this.hidePanel();
    }, PANEL_IDLE_MS);
  }

  renderPanel() {
    const s = this.getStatus();
    clear(this.panel);
    this.statusLine = el('div', { class: 'bo-status' });
    const rows = [this.statusLine];
    if (s.mode === 'rec') {
      rows.push(el('div', { class: 'bo-grid' },
        iconButton('star', '印を付ける', async () => { this.bumpIdle(); await this.onPin(); this.flashMsg('印を付けました'); }),
        iconButton('memo', 'メモ', () => { this.hide(); this.onMemo(); }),
        iconButton('moon', '暗くする', () => this.hidePanel())));
      rows.push(holdButton('長押しで終了', () => { this.onStop(); }, { className: 'bo-stop', enabled: this.holdToStop, hint: '長押しで終了' }));
      rows.push(el('p', { class: 'bo-note' }, '5秒間操作しないと、元の暗い画面に戻ります'));
    } else {
      rows.push(el('div', { class: 'bo-grid two' },
        iconButton('back', '通常の画面に戻る', () => this.exit()),
        iconButton('moon', '暗くする', () => this.hidePanel())));
      rows.push(el('p', { class: 'bo-note' }, '終わると、自動で通常の画面に戻ります'));
    }
    this.flashLine = el('p', { class: 'bo-flash', 'aria-live': 'polite' });
    rows.push(this.flashLine);
    this.panel.append(...rows);
    this.panel.addEventListener('pointerdown', () => this.bumpIdle());
    this.renderPanelStatus();
  }

  renderPanelStatus() {
    if (!this.statusLine) return;
    const s = this.getStatus();
    clear(this.statusLine);
    if (s.mode === 'rec') {
      this.statusLine.append(
        el('span', { class: 'bo-time' }, fmtTime(s.elapsed || 0)),
        el('span', { class: 'bo-saved' }, `端末に保存済み ${fmtTime(s.saved || 0)}`));
    } else {
      this.statusLine.append(
        el('span', { class: 'bo-proc' }, s.progressText || '処理中'),
        el('div', { class: 'bo-bar' }, el('div', { class: 'bo-bar-fill', style: `width:${Math.round((s.progressRatio || 0) * 100)}%` })));
    }
    const notes = (s.notes || []).slice(-1);
    if (notes.length) this.statusLine.appendChild(el('p', { class: 'bo-warnline' }, notes[0]));
  }

  flashMsg(text) {
    if (!this.flashLine) return;
    this.flashLine.textContent = text;
    clearTimeout(this.flashTimer);
    this.flashTimer = setTimeout(() => { if (this.flashLine) this.flashLine.textContent = ''; }, 1500);
  }

  /** 録音が止まったときなど：画面を明るくし、振動で知らせる */
  alert({ title, message, primary, onPrimary, secondary, onSecondary }) {
    this.closeAlert();
    try { if (navigator.vibrate) navigator.vibrate([400, 200, 400, 200, 400]); } catch { /* 無視 */ }
    const busy = (b, fn) => async () => {
      b.disabled = true;
      try { await fn(); } finally { b.disabled = false; }
    };
    const p = el('button', { class: 'bo-alert-primary', type: 'button' }, primary);
    p.addEventListener('click', busy(p, async () => { await onPrimary(); }));
    const s = secondary ? el('button', { class: 'bo-alert-secondary', type: 'button' }, secondary) : null;
    if (s) s.addEventListener('click', busy(s, async () => { await onSecondary(); }));
    const ico = document.createElement('span');
    ico.className = 'bo-alert-ico';
    ico.innerHTML = ICONS.warn;
    this.alertBox = el('div', { class: 'bo-alert', role: 'alertdialog', 'aria-modal': 'true', 'aria-labelledby': 'bo-alert-title' },
      el('div', { class: 'bo-alert-body' }, ico, el('h2', { id: 'bo-alert-title' }, title), el('p', {}, message)),
      el('div', { class: 'bo-alert-actions' }, p, s));
    document.body.appendChild(this.alertBox);
    p.focus();
  }

  closeAlert() {
    if (this.alertBox) { this.alertBox.remove(); this.alertBox = null; }
  }

  get alerting() { return !!this.alertBox; }
}
