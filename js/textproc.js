// 文字処理・ファイル出力・検証（ブラウザに依存しない純粋関数。Nodeで単体テスト可能）
import { fmtTime, parseTimeTag, fmtDateTime } from './util.js';

// ---------------------------------------------------------------------------
// 議事録テンプレート（D-12）
// ---------------------------------------------------------------------------
export const TEMPLATES = [
  { id: 'standard', name: '定例会議', focus: '進捗状況、決定事項、課題、次回までのアクション', sections: ['進捗状況', '議論の内容', '課題・リスク'] },
  { id: 'client', name: '商談・ヒアリング', focus: '相手方の課題・要望、こちらの提案内容、合意事項、次のアクション', sections: ['相手方の課題・要望', '提案・説明した内容', '相手方の反応・懸念'] },
  { id: 'legal', name: '案件打合せ（法務）', focus: '事実関係、論点、検討結果・方針、依頼事項と期限、留意点（正確性を最優先）', sections: ['事実関係', '論点', '検討結果・方針', '留意点'] },
  { id: 'brainstorm', name: 'ブレスト・企画', focus: 'アイデアの一覧、評価（利点・懸念）、優先順位', sections: ['出たアイデア', '評価（利点・懸念）', '優先順位'] },
  { id: 'oneonone', name: '1on1・面談', focus: '話題、本人の状況・懸念、合意事項、フォローアップ', sections: ['話題', '状況・懸念', 'フォローアップ'] },
  { id: 'trouble', name: '課題対策会議', focus: '事象、原因、暫定対策と恒久対策、担当と期限', sections: ['事象', '原因', '対策（暫定・恒久）'] },
];
export const templateById = (id) => TEMPLATES.find((t) => t.id === id) || TEMPLATES[0];

// ---------------------------------------------------------------------------
// ケバ取り（C-11）。元の文は別に保存し、ここでは表示・出力用の文を作る。
// ---------------------------------------------------------------------------
const FILLER_WORDS = [
  'えーっと', 'えーと', 'ええと', 'えっと', 'えー+', 'あのー+', 'あのぅ', 'そのー+', 'うーん+', 'うーむ', 'んー+', 'あー+',
  'あの(?=、)', 'その(?=、)', 'まあ(?=、)', 'なんか(?=、)',
];
const FILLER_RE = new RegExp(`(^|[、。！？!?\\s])(?:${FILLER_WORDS.join('|')})(?:[、,]\\s*|\\s+|(?=[。！？!?]|$))`, 'g');

export function removeFillers(text) {
  if (!text) return text;
  let prev;
  let out = String(text);
  let guard = 0;
  do { prev = out; out = out.replace(FILLER_RE, '$1'); guard++; } while (out !== prev && guard < 10);
  out = out.replace(/、{2,}/g, '、').replace(/^[、。\s]+/, '').replace(/\s{2,}/g, ' ').trim();
  return out || String(text).trim();
}

// ---------------------------------------------------------------------------
// 書き起こし
// segment: { id, start, end, speaker, label, text, raw, estimated }
// ---------------------------------------------------------------------------
export function speakerLetter(i) {
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  return i < letters.length ? `話者${letters[i]}` : `話者${i + 1}`;
}

/** 音声を最長 chunkSec ごとに区切る。区切り位置は前後30秒で最も静かな秒を選ぶ */
export function planChunks(durationSec, chunkSec = 1200, levels = [], searchSec = 30) {
  const chunks = [];
  if (!(durationSec > 0)) return chunks;
  let start = 0;
  while (durationSec - start > chunkSec) {
    const target = start + chunkSec;
    let cut = target;
    let best = Infinity;
    const lo = Math.max(start + chunkSec - searchSec, start + 60);
    const hi = Math.min(start + chunkSec, durationSec - 1);   // 上限を超えない側で探す
    for (let t = Math.floor(lo); t <= Math.floor(hi); t++) {
      const v = Number.isFinite(levels[t]) ? levels[t] : 0;
      if (v <= best) { best = v; cut = t; } // 同じ静かさなら後ろ（区切りが長くなる側）を選ぶ
    }
    if (cut <= start) cut = target;
    chunks.push({ start, end: cut });
    start = cut;
  }
  const rest = durationSec - start;
  if (chunks.length && rest < 60 && (chunks[chunks.length - 1].end - chunks[chunks.length - 1].start) + rest <= chunkSec + searchSec) {
    chunks[chunks.length - 1].end = durationSec;
  } else {
    chunks.push({ start, end: durationSec });
  }
  return chunks;
}

/** 区切りごとの話者ラベル（c0:spk_1 など）に、全体を通した表示名を割り当てる */
export function assignSpeakers(segments, mapping = {}) {
  const order = [];
  const globalOf = (seg) => {
    const m = mapping[seg.label];
    if (m && m.speaker) return m.speaker;
    // 対応づけがない場合：区切りをまたいでも同じ番号は同じ人と仮定
    const num = String(seg.label || '').split(':').pop();
    return `auto:${num}`;
  };
  for (const s of segments) {
    const g = globalOf(s);
    if (!order.includes(g)) order.push(g);
  }
  const nameOf = {};
  order.forEach((g, i) => {
    const named = Object.values(mapping).find((m) => m && m.speaker === g && m.name);
    nameOf[g] = named ? named.name : speakerLetter(i);
  });
  return segments.map((s) => ({ ...s, speaker: nameOf[globalOf(s)] }));
}

/** AIの修正案を適用する。極端な長さの変化（言い換え・創作）は採用しない */
export function applyCorrections(segments, corrections = []) {
  const byId = new Map();
  for (const c of corrections || []) {
    if (c && Number.isInteger(c.id) && typeof c.text === 'string' && c.text.trim()) byId.set(c.id, c.text.trim());
  }
  let applied = 0;
  const out = segments.map((s) => {
    const t = byId.get(s.id);
    if (!t || t === s.text) return s;
    const ratio = t.length / Math.max(1, s.text.length);
    if (ratio < 0.5 || ratio > 2) return s;
    applied++;
    return { ...s, text: t };
  });
  return { segments: out, applied };
}

/** 同じ話者の短い発言が続く場合にまとめる */
export function mergeSegments(segments, gapSec = 1.5, maxLen = 240) {
  const out = [];
  for (const s of segments) {
    const last = out[out.length - 1];
    if (last && last.label === s.label && s.start != null && last.end != null &&
        s.start - last.end <= gapSec && (last.text.length + s.text.length) <= maxLen) {
      last.text = joinText(last.text, s.text);
      last.raw = joinText(last.raw || last.text, s.raw || s.text);
      last.end = s.end != null ? s.end : last.end;
    } else {
      out.push({ ...s });
    }
  }
  return out.map((s, i) => ({ ...s, id: i }));
}

function joinText(a, b) {
  if (!a) return b;
  if (!b) return a;
  const ascii = /[A-Za-z0-9]$/.test(a) && /^[A-Za-z0-9]/.test(b);
  return ascii ? `${a} ${b}` : `${a}${b}`;
}

export function segmentsToTxt(segments, meta = {}) {
  const lines = [];
  if (meta.title) lines.push(`# ${meta.title}`);
  if (meta.startedAt) lines.push(`# 録音開始: ${fmtDateTime(meta.startedAt)}`);
  if (lines.length) lines.push('');
  for (const s of segments) {
    lines.push(`[${fmtTime(s.start || 0)}] ${s.speaker || '話者'}: ${s.text}`);
  }
  return lines.join('\r\n') + '\r\n';
}

const srtTime = (sec) => {
  const ms = Math.max(0, Math.round((sec || 0) * 1000));
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  const s = Math.floor((ms % 60000) / 1000);
  const r = ms % 1000;
  const p = (n, w = 2) => String(n).padStart(w, '0');
  return `${p(h)}:${p(m)}:${p(s)},${p(r, 3)}`;
};

export function segmentsToSrt(segments) {
  const out = [];
  segments.forEach((s, i) => {
    const next = segments[i + 1];
    const start = s.start || 0;
    let end = s.end != null ? s.end : (next ? next.start : start + 5);
    if (!(end > start)) end = start + 2;
    out.push(String(i + 1), `${srtTime(start)} --> ${srtTime(end)}`, `${s.speaker || ''}: ${s.text}`, '');
  });
  return out.join('\r\n');
}

/** 発言の前後 tolSec 秒以内のメモ */
export function memosNear(segment, memos, tolSec = 12) {
  const a = (segment.start || 0) - tolSec;
  const b = (segment.end != null ? segment.end : segment.start || 0) + tolSec;
  return memos.filter((m) => Number.isFinite(m.offset) && m.offset >= a && m.offset <= b);
}

// ---------------------------------------------------------------------------
// メモ
// memo: { id, deviceId, deviceName, t(作成時刻ms), offset(秒), text, kind, updatedAt, deleted }
// ---------------------------------------------------------------------------
export function memoOffset(memo, startedAt) {
  if (Number.isFinite(memo.offset)) return memo.offset;
  if (Number.isFinite(memo.t) && Number.isFinite(startedAt)) return Math.round((memo.t - startedAt) / 100) / 10;
  return null;
}

/** 複数端末のメモを統合（同じIDは更新が新しい方を採用、削除済みは除く、時刻順） */
export function mergeMemos(...lists) {
  const map = new Map();
  for (const list of lists) {
    if (!Array.isArray(list)) continue;
    for (const m of list) {
      if (!m || !m.id) continue;
      const prev = map.get(m.id);
      if (!prev || (m.updatedAt || 0) >= (prev.updatedAt || 0)) map.set(m.id, m);
    }
  }
  return Array.from(map.values())
    .filter((m) => !m.deleted)
    .sort((a, b) => (a.offset ?? Infinity) - (b.offset ?? Infinity) || (a.t || 0) - (b.t || 0));
}

export function memosToTxt(memos, meta = {}) {
  const lines = [];
  if (meta.title) lines.push(`# ${meta.title} メモ`, '');
  for (const m of memos) {
    const time = Number.isFinite(m.offset) ? (m.offset < 0 ? '開始前' : fmtTime(m.offset)) : '時刻なし';
    const mark = m.kind === 'star' ? '★' : '';
    const dev = m.deviceName ? `（${m.deviceName}）` : '';
    lines.push(`[${time}] ${mark}${m.text}${dev}`);
  }
  return lines.join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------
// 議事録（構造化データ → Markdown）
// ---------------------------------------------------------------------------
function evid(item) {
  const tags = (item.evidence || []).map((e) => parseTimeTag(String(e))).filter((x) => x != null);
  const uniq = [...new Set(tags)].sort((a, b) => a - b).slice(0, 4);
  const s = uniq.map((t) => `[${fmtTime(t)}]`).join('');
  const warn = item.unverified ? ' ⚠根拠未確認' : '';
  const src = item.source === 'memo' ? '（メモより）' : item.source === 'material' ? '（資料より）' : '';
  return `${src}${s ? ' ' + s : ''}${warn}`;
}

const mdCell = (s) => String(s ?? '').replace(/\|/g, '｜').replace(/\r?\n/g, ' ');

export function minutesToMarkdown(min, meta = {}) {
  const L = [];
  L.push(`# ${meta.title || '会議'}`, '');
  const when = meta.startedAt ? `${fmtDateTime(meta.startedAt)}${meta.endedAt ? '〜' + fmtDateTime(meta.endedAt).slice(11) : ''}` : '';
  if (when) L.push(`- 日時: ${when}${meta.durationSec ? `（約${Math.max(1, Math.round(meta.durationSec / 60))}分）` : ''}`);
  const people = (meta.participants && meta.participants.length ? meta.participants : min.participants) || [];
  if (people.length) L.push(`- 参加者: ${people.join('、')}`);
  if (meta.projectName) L.push(`- 案件: ${meta.projectName}`);
  if (meta.templateName) L.push(`- 種類: ${meta.templateName}`);
  if (meta.folderUrl) L.push(`- 保存先: [OneDriveのフォルダ](${meta.folderUrl})`);
  L.push('');
  L.push('## 要約', '', (min.summary || '（要約なし）').trim(), '');

  L.push('## 決定事項', '');
  if (min.decisions && min.decisions.length) min.decisions.forEach((d) => L.push(`- ${d.text}${evid(d)}`));
  else L.push('- （決定事項はありません）');
  L.push('');

  L.push('## ToDo', '');
  if (min.todos && min.todos.length) {
    L.push('| 担当 | 内容 | 期限 | 根拠 |', '|---|---|---|---|');
    min.todos.forEach((t) => L.push(`| ${mdCell(t.owner || '未定')} | ${mdCell(t.task)} | ${mdCell(t.due || '未定')} | ${mdCell(evid(t).trim())} |`));
  } else L.push('- （ToDoはありません）');
  L.push('');

  for (const sec of min.sections || []) {
    if (!sec || !sec.heading) continue;
    L.push(`## ${sec.heading}`, '');
    const items = sec.items || [];
    if (items.length) items.forEach((it) => L.push(`- ${it.text}${evid(it)}`));
    else L.push('- （該当なし）');
    L.push('');
  }

  if (min.openIssues && min.openIssues.length) {
    L.push('## 未決事項・持ち越し', '');
    min.openIssues.forEach((o) => L.push(`- ${o.text}${evid(o)}`));
    L.push('');
  }
  if (min.carryover && min.carryover.length) {
    L.push('## 前回からの持ち越しToDoの状況', '');
    min.carryover.forEach((c) => L.push(`- ${c.task}（${c.owner || '担当未定'}）: ${c.status || '状況不明'}${evid(c)}`));
    L.push('');
  }
  if (min.nextMeeting) L.push('## 次回', '', min.nextMeeting, '');
  if (meta.materials && meta.materials.length) {
    L.push('## 参照した資料', '');
    meta.materials.forEach((n) => L.push(`- ${n}`));
    L.push('');
  }
  L.push('---', '', '※この議事録はAIが録音・メモ・資料から作成しました。[MM:SS] は録音内の根拠の位置です。「⚠根拠未確認」の項目は、書き起こしに対応する発言が見つからなかったものです。');
  return L.join('\n') + '\n';
}

/** 根拠の時刻を書き起こしと照合する（D-14）。各項目に unverified を付けて返す */
export function verifyEvidence(min, segments, tolSec = 15) {
  const has = (t) => segments.some((s) => {
    const a = (s.start || 0) - tolSec;
    const b = (s.end != null ? s.end : (s.start || 0)) + tolSec;
    return t >= a && t <= b;
  });
  let flagged = 0;
  const mark = (it) => {
    if (!it || typeof it !== 'object') return it;
    const times = (it.evidence || []).map((e) => parseTimeTag(String(e))).filter((x) => x != null);
    const fromOther = it.source === 'memo' || it.source === 'material';
    const ok = times.some(has);
    const unverified = !fromOther && !ok;
    if (unverified) flagged++;
    return { ...it, unverified };
  };
  const out = {
    ...min,
    decisions: (min.decisions || []).map(mark),
    todos: (min.todos || []).map(mark),
    openIssues: (min.openIssues || []).map(mark),
    sections: (min.sections || []).map((s) => ({ ...s, items: (s.items || []).map(mark) })),
    carryover: (min.carryover || []).map((c) => ({ ...c })),
  };
  return { minutes: out, flagged };
}

export function todosToCsv(todos, meta = {}) {
  const esc = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const rows = [['会議', '日付', '担当', '内容', '期限', '根拠（録音内の時刻）', '状態']];
  const date = meta.startedAt ? fmtDateTime(meta.startedAt).slice(0, 10) : '';
  for (const t of todos || []) {
    rows.push([meta.title || '', date, t.owner || '未定', t.task || '', t.due || '未定',
      (t.evidence || []).join(' '), '未完了']);
  }
  return '\uFEFF' + rows.map((r) => r.map(esc).join(',')).join('\r\n') + '\r\n';
}

// ---------------------------------------------------------------------------
// iCalendar（.ics）。行は75オクテットで折り返す（RFC 5545）
// ---------------------------------------------------------------------------
const icsEscape = (s) => String(s ?? '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
const icsTime = (ms) => new Date(ms).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');

export function foldIcsLine(line) {
  const enc = new TextEncoder();
  const out = [];
  let cur = '';
  let bytes = 0;
  for (const ch of line) {
    const b = enc.encode(ch).length;
    const limit = out.length === 0 ? 75 : 74; // 継続行は先頭の空白1オクテット分を引く
    if (bytes + b > limit) { out.push(cur); cur = ''; bytes = 0; }
    cur += ch; bytes += b;
  }
  out.push(cur);
  return out.join('\r\n ');
}

export function buildIcs(ev) {
  const lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//mtgrec//ja', 'CALSCALE:GREGORIAN', 'BEGIN:VEVENT',
    `UID:${ev.uid}`, `DTSTAMP:${icsTime(Date.now())}`,
    `DTSTART:${icsTime(ev.start)}`, `DTEND:${icsTime(ev.end)}`,
    `SUMMARY:${icsEscape(ev.title)}`,
  ];
  if (ev.description) lines.push(`DESCRIPTION:${icsEscape(ev.description)}`);
  if (ev.url) lines.push(`URL:${ev.url}`);
  lines.push('END:VEVENT', 'END:VCALENDAR');
  return lines.map(foldIcsLine).join('\r\n') + '\r\n';
}

/** Googleカレンダーの登録画面を開くURL（API連携なしの代替手段） */
export function calendarTemplateUrl(ev, account = '') {
  const p = new URLSearchParams({
    action: 'TEMPLATE',
    text: ev.title,
    dates: `${icsTime(ev.start)}/${icsTime(ev.end)}`,
    details: ev.description || '',
  });
  if (ev.url) p.set('location', ev.url);
  if (account) p.set('authuser', account);
  return `https://calendar.google.com/calendar/render?${p.toString()}`;
}

// ---------------------------------------------------------------------------
// 安全なMarkdown表示用の解析（HTMLは一切通さない）
// ---------------------------------------------------------------------------
export function parseMarkdown(md) {
  const blocks = [];
  const lines = String(md || '').replace(/\r\n/g, '\n').split('\n');
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }
    let m;
    if ((m = line.match(/^(#{1,3})\s+(.*)$/))) { blocks.push({ type: `h${m[1].length}`, text: m[2] }); i++; continue; }
    if (/^---+\s*$/.test(line)) { blocks.push({ type: 'hr' }); i++; continue; }
    if (/^\|.*\|\s*$/.test(line)) {
      const rows = [];
      while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) {
        const cells = lines[i].trim().slice(1, -1).split('|').map((c) => c.trim());
        if (!cells.every((c) => /^:?-{2,}:?$/.test(c))) rows.push(cells);
        i++;
      }
      blocks.push({ type: 'table', rows });
      continue;
    }
    if (/^\s*[-*]\s+/.test(line)) {
      const items = [];
      while (i < lines.length && /^\s*[-*]\s+/.test(lines[i])) { items.push(lines[i].replace(/^\s*[-*]\s+/, '')); i++; }
      blocks.push({ type: 'ul', items });
      continue;
    }
    const para = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,3}\s|\||\s*[-*]\s|---)/.test(lines[i])) { para.push(lines[i]); i++; }
    blocks.push({ type: 'p', text: para.join('\n') });
  }
  return blocks;
}

/** 行内の装飾：**太字**、[文字](https://…)、[MM:SS] */
export function parseInline(text) {
  const out = [];
  const re = /\*\*([^*]+)\*\*|\[([^\]]+)\]\((https:\/\/[^)\s]+)\)|\[((?:\d+:)?\d{1,3}:\d{2})\]/g;
  let last = 0;
  let m;
  const s = String(text || '');
  while ((m = re.exec(s))) {
    if (m.index > last) out.push({ type: 'text', text: s.slice(last, m.index) });
    if (m[1] != null) out.push({ type: 'bold', text: m[1] });
    else if (m[2] != null) out.push({ type: 'link', text: m[2], href: m[3] });
    else out.push({ type: 'time', text: `[${m[4]}]`, sec: parseTimeTag(m[4]) });
    last = re.lastIndex;
  }
  if (last < s.length) out.push({ type: 'text', text: s.slice(last) });
  return out;
}

// ---------------------------------------------------------------------------
// 行単位の差分（議事録の手直しから学ぶため）
// ---------------------------------------------------------------------------
export function lineDiff(beforeText, afterText, maxHunks = 30) {
  const a = String(beforeText || '').split(/\r?\n/);
  const b = String(afterText || '').split(/\r?\n/);
  const n = a.length;
  const m = b.length;
  if (n * m > 4_000_000) return [{ before: a.slice(0, 50), after: b.slice(0, 50) }];
  const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = a[i] === b[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }
  const hunks = [];
  let cur = null;
  let i = 0;
  let j = 0;
  const flush = () => { if (cur && (cur.before.length || cur.after.length)) hunks.push(cur); cur = null; };
  while (i < n || j < m) {
    if (i < n && j < m && a[i] === b[j]) { flush(); i++; j++; continue; }
    cur = cur || { before: [], after: [] };
    if (j < m && (i >= n || dp[i][j + 1] >= dp[i + 1][j])) { cur.after.push(b[j]); j++; } else { cur.before.push(a[i]); i++; }
  }
  flush();
  return hunks.filter((h) => (h.before.join('') + h.after.join('')).trim()).slice(0, maxHunks);
}

// ---------------------------------------------------------------------------
// 費用の推定（REQ-11）
// ---------------------------------------------------------------------------
export const DEFAULT_PRICES = [
  { model: 'gemini-3.5-transcribe', in: 2.0, out: 12.0 },
  { model: 'gemini-3.8-flash', in: 0.75, out: 3.75, until: '2026-12-31' },
  { model: 'gemini-3.8-flash', in: 1.5, out: 7.5, from: '2027-01-01' },
  { model: 'gemini-3.7-flash', in: 0.75, out: 3.75, until: '2026-12-31' },
  { model: 'gemini-3.7-flash', in: 1.5, out: 7.5, from: '2027-01-01' },
  { model: '*', in: 2.0, out: 12.0 }, // 不明なモデルは高めに見積もる
];

export function priceFor(model, at, prices = DEFAULT_PRICES) {
  const day = new Date(at).toISOString().slice(0, 10);
  const cands = prices.filter((p) => p.model === model && (!p.from || day >= p.from) && (!p.until || day <= p.until));
  return cands[0] || prices.find((p) => p.model === '*') || { in: 2, out: 12 };
}

export function costOf(rec, prices = DEFAULT_PRICES) {
  const p = priceFor(rec.model, rec.at || Date.now(), prices);
  const out = (rec.outTok || 0) + (rec.thoughtTok || 0);
  return ((rec.inTok || 0) * p.in + out * p.out) / 1e6;
}

export function monthlySummary(records, now = Date.now(), prices = DEFAULT_PRICES) {
  const d = new Date(now);
  const key = `${d.getFullYear()}-${d.getMonth()}`;
  let cost = 0;
  let calls = 0;
  let audioSec = 0;
  for (const r of records || []) {
    const rd = new Date(r.at);
    if (`${rd.getFullYear()}-${rd.getMonth()}` !== key) continue;
    cost += costOf(r, prices);
    calls++;
    audioSec += r.audioSec || 0;
  }
  return { cost, calls, audioSec };
}
