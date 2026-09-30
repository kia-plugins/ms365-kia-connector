/**
 * The calendar half of the ms365 pull (spec 2026-09-30 §6): calendars sync
 * after mail in the same pull; the item batches carry the old calendar
 * cursor and only the last batch commits the new one.
 */
import type { Batch } from '@kiagent/connector-sdk';
import { createMs365Source, SCOPES } from '../source';
import type { Ms365Cursor } from '../cursor';
import type { Ms365Item } from '../source';
import {
  graphFetch,
  graphMsg,
  instantClock,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365Item>;
const G = 'https://graph.microsoft.com/v1.0';
const INBOX_ROOT = [{ id: 'INBOX-ID', name: 'Inbox' }];
const live: Ms365Cursor = {
  v: 2,
  attachments: 1,
  phase: 'live',
  folders: { 'INBOX-ID': { delta: `${G}/inbox-d` } },
  pending: [],
  retry: [],
};
const mail: GraphWorld = {
  urls: {
    [`${G}/inbox-d`]: {
      value: [{ id: 'm1', conversationId: 'CM', parentFolderId: 'INBOX-ID' }],
      '@odata.deltaLink': `${G}/inbox-d2`,
    },
  },
  conversations: { CM: [graphMsg({ conversationId: 'CM', parentFolderId: 'INBOX-ID' })] },
};
const event = {
  id: 'EV1',
  iCalUId: 'UID1',
  subject: 'Standup',
  start: { dateTime: '2026-10-01T09:00:00.0000000' },
  end: { dateTime: '2026-10-01T09:15:00.0000000' },
};
const calendars = { value: [{ id: 'CAL', name: 'Calendar', isDefaultCalendar: true }] };

function run(world: GraphWorld, tenantKind: 'work' | 'personal' = 'work') {
  const { fetchFn } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  const { session, logs } = makeSession({
    config: { tenantKind, folderRoots: INBOX_ROOT },
  });
  const seen: B[] = [];
  const done = (async () => {
    for await (const b of source.pull(session, live)) seen.push(b as B);
  })();
  return { source, seen, logs, done };
}

test('one pull yields mail then calendar; the last batch commits the calendar cursor', async () => {
  const { source, seen, done } = run({ ...mail, calendars, calendarViews: { CAL: [event] } });
  await done;
  const all = seen.flatMap((b) => b.items);
  const cal = all.find((i) => 'calendarEvent' in i)!;
  expect(all.some((i) => 'conversationId' in i)).toBe(true);
  expect((source.toDocument(cal) as { type: string }).type).toBe('calendar.event');
  const last = seen[seen.length - 1];
  expect(last.cursor.calendar?.since).toBeDefined();
  expect(last.cursor.calendar?.cals.CAL).toEqual({ EV1: '2026-10-01T09:00:00.000Z' });
  // Every earlier batch still carries the prior (absent) calendar half.
  for (const b of seen.slice(0, -1)) expect(b.cursor.calendar).toBeUndefined();
});

test('no Calendars.Read consent yet (403, no calendar cursor): mail syncs, one warn, no calendar cursor', async () => {
  const { seen, logs, done } = run({ ...mail, calendarStatus: 403 });
  await done;
  expect(seen.flatMap((b) => b.items).some((i) => 'conversationId' in i)).toBe(true);
  const warns = logs.filter((l) => l.level === 'warn' && l.msg.includes('Calendars.Read'));
  expect(warns).toHaveLength(1);
  expect(seen[seen.length - 1].cursor.calendar).toBeUndefined();
});

test('a 500 on calendars throws after the mail batches were yielded', async () => {
  const { seen, done } = run({ ...mail, calendarStatus: 500 });
  await expect(done).rejects.toThrow();
  expect(seen.flatMap((b) => b.items).some((i) => 'conversationId' in i)).toBe(true);
});

test('descriptor declares calendar.event and the scopes ask for Calendars.Read', () => {
  const source = createMs365Source(makeHost(graphFetch({}).fetchFn), instantClock);
  expect(source.descriptor.documentTypes).toEqual(['email.thread', 'attachment', 'calendar.event']);
  expect(SCOPES).toContain('Calendars.Read');
});

test('a personal account links the event on outlook.live.com (mail does the same)', async () => {
  const { source, seen, done } = run({ ...mail, calendars, calendarViews: { CAL: [event] } }, 'personal');
  await done;
  const cal = seen.flatMap((b) => b.items).find((i) => 'calendarEvent' in i)!;
  expect((source.toDocument(cal) as { url: string }).url).toBe('https://outlook.live.com/calendar/item/EV1');
});

test('an aborted pull stops before the calendar half (pause, quit)', async () => {
  const { fetchFn, calls } = graphFetch({ ...mail, calendars, calendarViews: { CAL: [event] } });
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  const ac = new AbortController();
  const { session } = makeSession({ config: { tenantKind: 'work', folderRoots: INBOX_ROOT }, signal: ac.signal });
  ac.abort();
  for await (const _ of source.pull(session, live)) { /* drain */ }
  expect(calls.some((u) => u.includes('/me/calendars'))).toBe(false);
});
