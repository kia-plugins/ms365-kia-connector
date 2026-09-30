import { pullCalendars } from '../calendar/pull';

const DAY = 86_400_000;
const NOW = Date.parse('2026-10-01T00:00:00Z');
const cals = { value: [
  { id: 'C1', name: 'Calendar', hexColor: '#0078d4', isDefaultCalendar: true, canEdit: true },
  { id: 'C2', name: 'Birthdays', canEdit: false },
] };
const ev = (id: string, startIso: string) => ({
  id, iCalUId: `U${id}`, subject: id,
  start: { dateTime: startIso.replace('Z', '') }, end: { dateTime: startIso.replace('Z', '') },
});

function fake(views: Record<string, unknown[]>) {
  const urls: string[] = [];
  return {
    urls,
    request: async <T,>(url: string): Promise<T> => {
      urls.push(url);
      if (url.includes('/me/calendars?')) return cals as T;
      const m = /\/me\/calendars\/([^/]+)\/calendarView/.exec(url);
      return { value: views[m![1]] ?? [] } as T;
    },
  };
}

test('first pull: default + owned calendars, from 365 days back, cursor fixed', async () => {
  const c = fake({ C1: [ev('a', '2026-09-01T09:00:00Z')] });
  const r = await pullCalendars(c, {}, undefined, NOW);
  expect(r.items.map((i) => i.calendarEvent.id)).toEqual(['a']);
  expect(r.cursor.since).toBe(new Date(NOW - 365 * DAY).toISOString());
  expect(c.urls.some((u) => u.includes('/calendars/C2/'))).toBe(false); // not owned (canEdit false)
  expect(c.urls.find((u) => u.includes('/C1/calendarView'))).toContain(
    `startDateTime=${encodeURIComponent(new Date(NOW - 365 * DAY).toISOString())}`,
  );
});

test('later pull: sliding window; a vanished in-window event is a deletion; ageing out is not', async () => {
  const prior = {
    since: new Date(NOW - 365 * DAY).toISOString(),
    cals: { C1: { old: '2026-08-01T09:00:00.000Z', gone: '2026-09-20T09:00:00.000Z', kept: '2026-09-25T09:00:00.000Z' } },
  };
  const c = fake({ C1: [ev('kept', '2026-09-25T09:00:00Z')] });
  const r = await pullCalendars(c, {}, prior, NOW);
  expect(r.deletions).toEqual([{ externalId: 'C1:gone', type: 'calendar.event' }]);
  expect(Object.keys(r.cursor.cals.C1)).toEqual(['kept']); // 'old' left the window: dropped, not deleted
  expect(r.cursor.since).toBe(prior.since);
  expect(c.urls.find((u) => u.includes('/C1/calendarView'))).toContain(
    `startDateTime=${encodeURIComponent(new Date(NOW - 35 * DAY).toISOString())}`,
  );
});

test('explicit calendarRoots win; an unticked calendar is not listed', async () => {
  const c = fake({ C2: [ev('b', '2026-09-30T09:00:00Z')] });
  const r = await pullCalendars(c, { calendarRoots: [{ id: 'C2', name: 'Birthdays' }] }, undefined, NOW);
  expect(r.items.map((i) => i.calendar.id)).toEqual(['C2']);
  const none = await pullCalendars(fake({}), { calendarRoots: [] }, undefined, NOW);
  expect(none.items).toEqual([]);
});

test('a calendar colour that is not #rrggbb (Graph "auto", "") falls back to the page palette', async () => {
  const c = {
    request: async <T,>(url: string): Promise<T> =>
      (url.includes('/me/calendars?')
        ? { value: [
            { id: 'A', name: 'Auto', hexColor: 'auto', isDefaultCalendar: true },
            { id: 'E', name: 'Empty', hexColor: '', canEdit: true },
            { id: 'H', name: 'Hex', hexColor: '#A1B2C3', canEdit: true },
          ] }
        : { value: [ev('x', '2026-09-30T09:00:00Z')] }) as T,
  };
  const r = await pullCalendars(c, {}, undefined, NOW);
  expect(r.items.map((i) => i.calendar.color)).toEqual([null, null, '#A1B2C3']);
});

test('calendarView asks for UTC times and text bodies', async () => {
  const seen: Array<Record<string, string> | undefined> = [];
  const c = {
    request: async <T,>(url: string, opts?: { extraHeaders?: Record<string, string> }): Promise<T> => {
      if (url.includes('/calendarView')) seen.push(opts?.extraHeaders);
      return (url.includes('/me/calendars?') ? cals : { value: [] }) as T;
    },
  };
  await pullCalendars(c, {}, undefined, NOW);
  expect(seen[0]).toEqual({ prefer: 'outlook.timezone="UTC", outlook.body-content-type="text"' });
});
