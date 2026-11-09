// 利用履歴・改善記録（REQ-11）
import { kvGet, kvSet } from './db.js';

const MAX_USAGE = 4000;
const MAX_LOG = 500;

export async function logUsage(rec) {
  const list = (await kvGet('usage', [])) || [];
  list.push({ at: Date.now(), ...rec });
  if (list.length > MAX_USAGE) list.splice(0, list.length - MAX_USAGE);
  await kvSet('usage', list);
}

export async function getUsage() { return (await kvGet('usage', [])) || []; }

/** 改善記録（会議の処理結果、単語帳への追加、手直しからの学習など） */
export async function logKaizen(type, text, extra = {}) {
  const list = (await kvGet('kaizen', [])) || [];
  list.unshift({ at: Date.now(), type, text, ...extra });
  if (list.length > MAX_LOG) list.length = MAX_LOG;
  await kvSet('kaizen', list);
}

export async function getKaizen() { return (await kvGet('kaizen', [])) || []; }
