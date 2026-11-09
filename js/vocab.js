// 案件別の単語帳・話者・書き方ルール（REQ-7 / C-10）。OneDrive経由で2台に共有する。
import { kvGet, kvSet } from './db.js';
import { uuid, debounce } from './util.js';
import { logKaizen } from './usage.js';

export const COMMON_ID = 'common';
let vocab = null;
let uploader = null; // 起動時に OneDrive への保存処理を登録する
const listeners = new Set();

function emptyProject(id, name, description = '') {
  return { id, name, description, terms: [], speakers: [], rules: [], updatedAt: Date.now() };
}

export async function loadVocab() {
  vocab = (await kvGet('vocab')) || { version: 1, updatedAt: 0, projects: [] };
  if (!vocab.projects.find((p) => p.id === COMMON_ID)) {
    vocab.projects.unshift(emptyProject(COMMON_ID, '全案件共通', 'すべての会議に使う用語・人名'));
    await kvSet('vocab', vocab);
  }
  return vocab;
}

export const getVocab = () => vocab;
export const onVocabChange = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };
export function setVocabUploader(fn) { uploader = fn; }

const scheduleUpload = debounce(() => { if (uploader) uploader(vocab).catch((e) => console.warn('単語帳の共有に失敗', e)); }, 3000);

export async function saveVocab({ sync = true } = {}) {
  vocab.updatedAt = Date.now();
  await kvSet('vocab', vocab);
  for (const fn of listeners) { try { fn(vocab); } catch (e) { console.error(e); } }
  if (sync) scheduleUpload();
}

export function projects() { return vocab.projects.filter((p) => !p.deleted); }
export function projectById(id) { return vocab.projects.find((p) => p.id === id && !p.deleted) || vocab.projects.find((p) => p.id === COMMON_ID); }

export async function addProject(name, description = '') {
  const p = emptyProject(`p-${uuid().slice(0, 8)}`, name.trim(), description);
  vocab.projects.push(p);
  await saveVocab();
  return p;
}

export async function renameProject(id, name) {
  const p = projectById(id);
  if (!p || p.id === COMMON_ID) return;
  p.name = name.trim();
  p.updatedAt = Date.now();
  await saveVocab();
}

export async function deleteProject(id) {
  if (id === COMMON_ID) return;
  const p = vocab.projects.find((x) => x.id === id);
  if (p) { p.deleted = true; p.updatedAt = Date.now(); await saveVocab(); }
}

const live = (arr) => (arr || []).filter((x) => !x.deleted);

/** その案件で使う用語（共通＋案件） */
export function termsFor(projectId) {
  const common = projectById(COMMON_ID);
  const p = projectId && projectId !== COMMON_ID ? projectById(projectId) : null;
  const all = [...live(common && common.terms), ...live(p && p.terms)];
  const seen = new Set();
  return all.filter((t) => { const k = `${t.correct}|${t.wrong || ''}`; if (seen.has(k)) return false; seen.add(k); return true; });
}

export function rulesFor(projectId) {
  const common = projectById(COMMON_ID);
  const p = projectId && projectId !== COMMON_ID ? projectById(projectId) : null;
  return [...live(common && common.rules), ...live(p && p.rules)].map((r) => r.text);
}

export function speakersFor(projectId) {
  const p = projectById(projectId);
  return live(p && p.speakers);
}

const norm = (s) => String(s || '').trim();

/** 用語を追加（重複は追加しない）。追加した用語を返す */
export async function addTerms(projectId, terms, source = '手入力') {
  const p = projectById(projectId);
  const existing = new Set(live(p.terms).map((t) => `${norm(t.correct)}|${norm(t.wrong)}`));
  const added = [];
  for (const t of terms || []) {
    const correct = norm(t.correct);
    const wrong = norm(t.wrong);
    if (!correct || correct.length > 60 || wrong === correct) continue;
    const key = `${correct}|${wrong}`;
    if (existing.has(key)) continue;
    existing.add(key);
    const rec = { id: uuid().slice(0, 12), correct, wrong, category: norm(t.category) || '', source, addedAt: Date.now(), updatedAt: Date.now() };
    p.terms.push(rec);
    added.push(rec);
  }
  if (added.length) {
    p.updatedAt = Date.now();
    await saveVocab();
    if (source !== '手入力') await logKaizen('vocab', `「${p.name}」に${added.length}語を自動追加（${source}）：${added.slice(0, 5).map((a) => a.correct).join('、')}${added.length > 5 ? ' ほか' : ''}`);
  }
  return added;
}

export async function removeTerm(projectId, termId) {
  const p = projectById(projectId);
  const t = p.terms.find((x) => x.id === termId);
  if (t) { t.deleted = true; t.updatedAt = Date.now(); p.updatedAt = Date.now(); await saveVocab(); }
}

export async function addRules(projectId, rules, source = '手入力') {
  const p = projectById(projectId);
  const existing = new Set(live(p.rules).map((r) => norm(r.text)));
  const added = [];
  for (const text of rules || []) {
    const t = norm(text);
    if (!t || t.length > 120 || existing.has(t)) continue;
    existing.add(t);
    const rec = { id: uuid().slice(0, 12), text: t, source, addedAt: Date.now(), updatedAt: Date.now() };
    p.rules.push(rec);
    added.push(rec);
  }
  if (added.length) {
    p.updatedAt = Date.now();
    await saveVocab();
    if (source !== '手入力') await logKaizen('rule', `「${p.name}」に書き方のルールを${added.length}件追加：${added.map((a) => a.text).join(' / ')}`);
  }
  return added;
}

export async function removeRule(projectId, ruleId) {
  const p = projectById(projectId);
  const r = p.rules.find((x) => x.id === ruleId);
  if (r) { r.deleted = true; r.updatedAt = Date.now(); await saveVocab(); }
}

/** 話者の名前・役割を記憶（C-10） */
export async function rememberSpeakers(projectId, names, role = '') {
  const p = projectById(projectId);
  let changed = false;
  for (const name of names || []) {
    const n = norm(name);
    if (!n || /^話者[A-Z0-9]+$/.test(n)) continue;
    if (live(p.speakers).some((s) => s.name === n)) continue;
    p.speakers.push({ id: uuid().slice(0, 12), name: n, role, updatedAt: Date.now() });
    changed = true;
  }
  if (changed) { p.updatedAt = Date.now(); await saveVocab(); }
}

export async function updateSpeaker(projectId, speakerId, patch) {
  const p = projectById(projectId);
  const s = p.speakers.find((x) => x.id === speakerId);
  if (s) { Object.assign(s, patch, { updatedAt: Date.now() }); await saveVocab(); }
}

/** 2台の単語帳を統合（項目ごとに更新が新しい方を採用。削除も記録として残す） */
export function mergeVocabData(a, b) {
  if (!b || !b.projects) return a;
  if (!a || !a.projects) return b;
  const byId = new Map(a.projects.map((p) => [p.id, JSON.parse(JSON.stringify(p))]));
  for (const rp of b.projects) {
    const lp = byId.get(rp.id);
    if (!lp) { byId.set(rp.id, JSON.parse(JSON.stringify(rp))); continue; }
    const mergeList = (x, y) => {
      const m = new Map((x || []).map((i) => [i.id, i]));
      for (const i of y || []) {
        const cur = m.get(i.id);
        if (!cur || (i.updatedAt || 0) > (cur.updatedAt || 0)) m.set(i.id, i);
      }
      return Array.from(m.values());
    };
    const newer = (rp.updatedAt || 0) > (lp.updatedAt || 0) ? rp : lp;
    byId.set(rp.id, {
      ...lp,
      name: newer.name,
      description: newer.description,
      deleted: newer.deleted,
      terms: mergeList(lp.terms, rp.terms),
      speakers: mergeList(lp.speakers, rp.speakers),
      rules: mergeList(lp.rules, rp.rules),
      updatedAt: Math.max(lp.updatedAt || 0, rp.updatedAt || 0),
    });
  }
  const projectsMerged = Array.from(byId.values());
  projectsMerged.sort((x, y) => (x.id === COMMON_ID ? -1 : y.id === COMMON_ID ? 1 : 0));
  return { version: 1, updatedAt: Math.max(a.updatedAt || 0, b.updatedAt || 0), projects: projectsMerged };
}

/** OneDrive上の単語帳と統合して保存し直す */
export async function mergeRemote(remote) {
  if (!remote) return false;
  const merged = mergeVocabData(vocab, remote);
  const changed = JSON.stringify(merged.projects) !== JSON.stringify(vocab.projects);
  vocab = merged;
  await saveVocab({ sync: false });
  return changed;
}
