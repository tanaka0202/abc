// Googleカレンダー（REQ-9 / P-1）
import { googleGetToken } from './auth.js';
import { getSettings } from './settings.js';
import { AppError } from './util.js';

const API = 'https://www.googleapis.com/calendar/v3';

async function call(method, path, body) {
  const t = await googleGetToken();
  const r = await fetch(`${API}${path}`, {
    method,
    headers: { Authorization: `Bearer ${t}`, 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (r.status === 204) return null;
  const j = await r.json().catch(() => ({}));
  if (!r.ok) {
    const err = new AppError(`Googleカレンダーの操作に失敗しました（${(j.error && j.error.message) || r.status}）`);
    err.status = r.status;
    throw err;
  }
  return j;
}

/** 会議IDから決まるイベントID（2台が同時に登録しても重複しない） */
export function eventIdFor(sessionId) {
  return ('mtg' + String(sessionId).toLowerCase().replace(/[^0-9a-v]/g, '')).slice(0, 60);
}

/** 予定を作成、既にあれば更新する */
export async function upsertEvent(sessionId, ev) {
  const cal = encodeURIComponent(getSettings().google.calendarId || 'primary');
  const id = eventIdFor(sessionId);
  const body = {
    id,
    summary: ev.title,
    description: ev.description || '',
    location: ev.url || '',
    start: { dateTime: new Date(ev.start).toISOString() },
    end: { dateTime: new Date(ev.end).toISOString() },
    source: ev.url ? { title: 'OneDriveの保存フォルダ', url: ev.url } : undefined,
    reminders: { useDefault: false, overrides: [] },
    transparency: 'transparent',
  };
  try {
    return await call('POST', `/calendars/${cal}/events`, body);
  } catch (e) {
    if (e.status === 409) {
      const { id: _omit, ...patch } = body;
      return call('PATCH', `/calendars/${cal}/events/${id}`, patch);
    }
    throw e;
  }
}

/** 今の前後の予定（P-1：予定から会議名・参加者を取り込む） */
export async function eventsAround(date = new Date(), beforeMin = 120, afterMin = 120) {
  const cal = encodeURIComponent(getSettings().google.calendarId || 'primary');
  const p = new URLSearchParams({
    timeMin: new Date(date.getTime() - beforeMin * 60000).toISOString(),
    timeMax: new Date(date.getTime() + afterMin * 60000).toISOString(),
    singleEvents: 'true',
    orderBy: 'startTime',
    maxResults: '20',
  });
  const j = await call('GET', `/calendars/${cal}/events?${p}`);
  return (j.items || [])
    .filter((e) => e.start && e.start.dateTime && !String(e.id).startsWith('mtg'))
    .map((e) => ({
      id: e.id,
      title: e.summary || '（無題）',
      start: Date.parse(e.start.dateTime),
      end: Date.parse(e.end && e.end.dateTime),
      attendees: (e.attendees || []).filter((a) => !a.self && !a.resource).map((a) => a.displayName || a.email.split('@')[0]),
      description: e.description || '',
    }));
}
