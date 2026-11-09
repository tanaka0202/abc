// 画面部品（通知、長押しボタン、確認ダイアログ、安全なMarkdown表示）
import { el, clear } from './util.js';
import { parseMarkdown, parseInline } from './textproc.js';

let toastTimer = null;
export function toast(message, type = 'info', ms = 4000) {
  const box = document.getElementById('toast');
  if (!box) return;
  clear(box);
  box.className = `toast show ${type}`;
  box.appendChild(el('span', {}, message));
  box.setAttribute('role', type === 'error' ? 'alert' : 'status');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { box.className = 'toast'; }, ms);
}

/** 長押しで実行するボタン（A-3：誤タップ防止） */
export function holdButton(label, onHold, { holdMs = 1500, className = '', enabled = true, hint = '長押しで実行' } = {}) {
  const fill = el('span', { class: 'hold-fill' });
  const btn = el('button', { class: `btn hold ${className}`, type: 'button', title: hint, 'aria-label': `${label}（${hint}）` }, fill, el('span', { class: 'hold-label' }, label));
  if (!enabled || holdMs <= 0) {
    btn.addEventListener('click', () => onHold());
    return btn;
  }
  let timer = null;
  let started = 0;
  let raf = 0;
  const reset = () => {
    clearTimeout(timer); cancelAnimationFrame(raf); timer = null;
    fill.style.width = '0%';
    btn.classList.remove('holding');
  };
  const animate = () => {
    const r = Math.min(1, (performance.now() - started) / holdMs);
    fill.style.width = `${r * 100}%`;
    if (r < 1 && timer) raf = requestAnimationFrame(animate);
  };
  btn.addEventListener('pointerdown', (e) => {
    if (e.button !== undefined && e.button !== 0) return;
    started = performance.now();
    btn.classList.add('holding');
    try { btn.setPointerCapture(e.pointerId); } catch { /* 無視 */ }
    timer = setTimeout(() => { reset(); onHold(); }, holdMs);
    raf = requestAnimationFrame(animate);
  });
  ['pointerup', 'pointercancel', 'pointerleave'].forEach((ev) => btn.addEventListener(ev, () => { if (timer) { reset(); toast(hint, 'info', 1500); } }));
  btn.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.shiftKey) onHold(); }); // キーボード操作：Shift+Enter
  btn.addEventListener('contextmenu', (e) => e.preventDefault());
  return btn;
}

/** 確認ダイアログ */
export function confirmDialog(message, { ok = 'OK', cancel = 'キャンセル', danger = false } = {}) {
  return new Promise((resolve) => {
    const dlg = el('dialog', { class: 'dialog' },
      el('p', {}, message),
      el('div', { class: 'row end' },
        el('button', { class: 'btn ghost', type: 'button', onclick: () => { dlg.close(); resolve(false); } }, cancel),
        el('button', { class: `btn ${danger ? 'danger' : 'primary'}`, type: 'button', onclick: () => { dlg.close(); resolve(true); } }, ok)));
    dlg.addEventListener('close', () => { dlg.remove(); resolve(false); });
    document.body.appendChild(dlg);
    dlg.showModal();
  });
}

/** 入力欄つきダイアログ（fields: [{name,label,value}]） */
export function formDialog(title, fields, { ok = '保存' } = {}) {
  return new Promise((resolve) => {
    const inputs = fields.map((f) => el('input', { class: 'input', value: f.value || '', placeholder: f.placeholder || '', 'aria-label': f.label }));
    const dlg = el('dialog', { class: 'dialog' },
      el('h3', {}, title),
      ...fields.map((f, i) => el('label', { class: 'field' }, el('span', {}, f.label), inputs[i])),
      el('div', { class: 'row end' },
        el('button', { class: 'btn ghost', type: 'button', onclick: () => { dlg.close(); } }, 'キャンセル'),
        el('button', {
          class: 'btn primary', type: 'button',
          onclick: () => { const out = {}; fields.forEach((f, i) => { out[f.name] = inputs[i].value; }); dlg.close(); resolve(out); },
        }, ok)));
    dlg.addEventListener('close', () => { dlg.remove(); resolve(null); });
    document.body.appendChild(dlg);
    dlg.showModal();
  });
}

function inline(text, onTime) {
  return parseInline(text).map((t) => {
    if (t.type === 'bold') return el('strong', {}, t.text);
    if (t.type === 'link') return el('a', { href: t.href, target: '_blank', rel: 'noopener noreferrer' }, t.text);
    if (t.type === 'time') {
      return onTime ? el('button', { class: 'time-tag', type: 'button', onclick: () => onTime(t.sec) }, t.text) : el('span', { class: 'time-tag' }, t.text);
    }
    for (const [mark, tip] of [['⚠根拠未確認', '根拠の時刻の前後に発言が見つかりませんでした'], ['△内容要確認', '根拠の時刻の発言と内容がほとんど一致しませんでした']]) {
      if (t.text.includes(mark)) {
        const [a, ...rest] = t.text.split(mark);
        return [a, el('span', { class: 'warn-tag', title: tip }, mark), rest.join(mark)];
      }
    }
    return t.text;
  });
}

/** Markdownを安全に表示（HTMLは解釈しない）。[MM:SS] を押すとその時刻へ */
export function renderMarkdown(md, onTime) {
  const root = el('div', { class: 'md' });
  for (const b of parseMarkdown(md)) {
    if (b.type === 'h1' || b.type === 'h2' || b.type === 'h3') root.appendChild(el(b.type, {}, inline(b.text, onTime)));
    else if (b.type === 'hr') root.appendChild(el('hr'));
    else if (b.type === 'ul') root.appendChild(el('ul', {}, b.items.map((it) => el('li', {}, inline(it, onTime)))));
    else if (b.type === 'table') {
      const [head, ...rows] = b.rows;
      root.appendChild(el('div', { class: 'table-wrap' }, el('table', {},
        el('thead', {}, el('tr', {}, (head || []).map((c) => el('th', {}, inline(c, onTime))))),
        el('tbody', {}, rows.map((r) => el('tr', {}, r.map((c) => el('td', {}, inline(c, onTime)))))))));
    } else root.appendChild(el('p', {}, inline(b.text, onTime)));
  }
  return root;
}

export function chip(text, kind = '') { return el('span', { class: `chip ${kind}` }, text); }

export function progressBar(ratio) {
  const bar = el('div', { class: 'progress' }, el('div', { class: 'progress-fill' }));
  if (ratio == null) bar.classList.add('indeterminate');
  else bar.firstChild.style.width = `${Math.round(Math.max(0, Math.min(1, ratio)) * 100)}%`;
  return bar;
}

export function section(title, ...children) {
  return el('section', { class: 'card' }, title ? el('h2', {}, title) : null, ...children);
}
