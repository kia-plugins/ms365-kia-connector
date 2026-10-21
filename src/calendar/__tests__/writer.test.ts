import type { EventIntent } from '@kiagent/connector-sdk';
import { GRAPH_BASE } from '../../graph-api';
import { instantClock, makeHost } from '../../testing/harness';
import { createMs365CalendarWriter } from '../writer';
import { calendarToDocument } from '../document';

type Res = { status: number; body?: unknown };
type Call = { method: string; url: string; body?: string };

function graph(routes: Array<[string, RegExp, Res | ((c: Call) => Res)]>) {
  const calls: Call[] = [];
  const fetchFn = async (url: string, init: { method?: string; body?: string }) => {
    const c: Call = { method: init.method ?? 'GET', url, body: init.body as string | undefined };
    calls.push(c);
    const hit = routes.find(([m, rx]) => m === c.method && rx.test(url.replace(GRAPH_BASE, '')));
    const r = hit ? (typeof hit[2] === 'function' ? hit[2](c) : hit[2]) : { status: 599 };
    return {
      status: r.status, statusText: '', headers: {},
      body: new TextEncoder().encode(r.body === undefined ? '' : JSON.stringify(r.body)),
    };
  };
  const w = createMs365CalendarWriter(makeHost(fetchFn as never) as never, instantClock);
  const json = (c: Call) => JSON.parse(c.body as string);
  const posts = (rx: RegExp) => calls.filter((c) => c.method === 'POST' && rx.test(c.url));
  return { w, calls, json, posts };
}

const ctx = { credentials: { accessToken: 'T' } } as never;
const INTENT: EventIntent = {
  accountId: 'A' as never, idempotencyKey: 'key-1', calendarId: 'CAL1', title: 'Parent evening',
  start: '2026-10-22T17:00:00+02:00', end: '2026-10-22T18:30:00+02:00', allDay: false,
  timeZone: 'Europe/Berlin', location: 'School hall', description: 'Bring the form',
  attendees: [{ email: 'bob@x.com', name: 'Bob' }], conference: 'none', category: 'kids',
  kia: { draftId: 'D9', sourceDocumentId: null },
};
const CATS = [
  'GET', /^\/me\/outlook\/masterCategories/, { status: 200, body: { value: [{ displayName: 'Work', color: 'preset0' }] } },
] as [string, RegExp, Res];
const CAL = (providers: string[]) =>
  ['GET', /^\/me\/calendars\/CAL1\?/, { status: 200, body: { id: 'CAL1', allowedOnlineMeetingProviders: providers } }] as [string, RegExp, Res];
const created = { id: 'EV1', iCalUId: 'UID1', webLink: 'https://outlook.office.com/x', onlineMeeting: null };
const CREATE = ['POST', /^\/me\/calendars\/CAL1\/events$/, { status: 201, body: created }] as [string, RegExp, Res];
const catsPost = ['POST', /^\/me\/outlook\/masterCategories$/, { status: 201, body: {} }] as [string, RegExp, Res];

test('create posts transactionId, categories, stamp and attendees', async () => {
  const { w, posts, json } = graph([CATS, catsPost, CAL([]), CREATE]);
  const r = await w.createEvent(INTENT, ctx);
  expect(r).toEqual({
    eventId: 'EV1', iCalUID: 'UID1', htmlLink: 'https://outlook.office.com/x',
    conferenceUrl: null, conferencePending: false, note: null,
  });
  const body = json(posts(/events$/)[0]);
  expect(body).toMatchObject({
    subject: 'Parent evening',
    body: { contentType: 'text', content: 'Bring the form' },
    start: { dateTime: '2026-10-22T17:00:00', timeZone: 'Europe/Berlin' },
    end: { dateTime: '2026-10-22T18:30:00', timeZone: 'Europe/Berlin' },
    location: { displayName: 'School hall' },
    attendees: [{ emailAddress: { address: 'bob@x.com', name: 'Bob' }, type: 'required' }],
    categories: ['Kids'],
    singleValueExtendedProperties: [
      { id: 'String {f5d1d8b2-3c4e-4a7b-9d61-6b1c2a9e7f10} Name kiaStamp', value: 'kids|D9' },
    ],
  });
  expect(body.transactionId).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-5[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  expect(body.isOnlineMeeting).toBeUndefined();
});

test('all-day events use midnight dates with the exclusive end', async () => {
  const { w, posts, json } = graph([CATS, catsPost, CAL([]), CREATE]);
  await w.createEvent({ ...INTENT, allDay: true, start: '2026-10-22', end: '2026-10-23', location: null, description: null }, ctx);
  const body = json(posts(/events$/)[0]);
  expect(body.isAllDay).toBe(true);
  expect(body.start.dateTime).toBe('2026-10-22T00:00:00');
  expect(body.end.dateTime).toBe('2026-10-23T00:00:00');
  expect(body.location).toBeUndefined();
  expect(body.body.content).toBe('');
});

test('same key gives the same transactionId', async () => {
  const a = graph([CATS, catsPost, CAL([]), CREATE]);
  const b = graph([CATS, catsPost, CAL([]), CREATE]);
  await a.w.createEvent(INTENT, ctx);
  await b.w.createEvent(INTENT, ctx);
  await b.w.createEvent({ ...INTENT, idempotencyKey: 'key-2' }, ctx);
  const ids = [a.posts(/events$/)[0], b.posts(/events$/)[0], b.posts(/events$/)[1]].map((c) => a.json(c).transactionId);
  expect(ids[0]).toBe(ids[1]);
  expect(ids[2]).not.toBe(ids[0]);
});

test('missing master categories are created once with preset colours; existing ones untouched', async () => {
  const { w, posts, json, calls } = graph([CATS, catsPost, CAL([]), CREATE]);
  await w.createEvent(INTENT, ctx);
  await w.createEvent({ ...INTENT, idempotencyKey: 'k2' }, ctx);
  const made = posts(/masterCategories$/).map((c) => json(c));
  expect(made).toEqual([
    { displayName: 'Personal', color: 'preset4' },
    { displayName: 'Kids', color: 'preset1' },
  ]);
  expect(calls.filter((c) => c.method === 'PATCH')).toHaveLength(0);
  expect(calls.filter((c) => /masterCategories/.test(c.url) && c.method === 'GET')).toHaveLength(1);
});

test('a category failure is only logged; the event is still created', async () => {
  const { w } = graph([
    ['GET', /^\/me\/outlook\/masterCategories/, { status: 403, body: { error: { code: 'AccessDenied' } } }],
    CAL([]), CREATE,
  ]);
  await expect(w.createEvent(INTENT, ctx)).resolves.toMatchObject({ eventId: 'EV1' });
});

test('teams only when allowed; otherwise note', async () => {
  const ok = graph([CATS, catsPost, CAL(['skypeForBusiness', 'teamsForBusiness']),
    ['POST', /events$/, { status: 201, body: { ...created, onlineMeeting: { joinUrl: 'https://teams/j' } } }]]);
  const r = await ok.w.createEvent({ ...INTENT, conference: 'auto' }, ctx);
  expect(ok.json(ok.posts(/events$/)[0])).toMatchObject({ isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' });
  expect(r.conferenceUrl).toBe('https://teams/j');
  expect(r.note).toBeNull();

  const no = graph([CATS, catsPost, CAL(['skypeForConsumer']), CREATE]);
  const r2 = await no.w.createEvent({ ...INTENT, conference: 'auto' }, ctx);
  expect(no.json(no.posts(/events$/)[0]).isOnlineMeeting).toBeUndefined();
  expect(r2.note).toBe("no video link: this Outlook account can't create Teams meetings");
});

test('403 uses the calendar reconnect copy, not the send copy', async () => {
  const { w } = graph([
    ['GET', /^\/me\?/, { status: 200, body: { mail: 'me@x.com' } }],
    CATS, catsPost, CAL([]),
    ['POST', /events$/, { status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'nope' } } }],
  ]);
  const err = (await w.createEvent(INTENT, ctx).catch((e) => e)) as Error;
  expect(err.message).toBe('reconnect Microsoft 365 (me@x.com) in Settings to allow adding events');
  expect(err.message).not.toMatch(/send permission/);
});

test('429 and other 4xx map to their own copy', async () => {
  const r429 = graph([CATS, catsPost, CAL([]), ['POST', /events$/, { status: 429, body: {} }]]);
  await expect(r429.w.createEvent(INTENT, ctx)).rejects.toThrow('rate-limited: Outlook is busy — try again in a minute');
  const r400 = graph([CATS, catsPost, CAL([]),
    ['POST', /events$/, { status: 400, body: { error: { code: 'ErrorInvalidRequest', message: 'bad start' } } }]]);
  await expect(r400.w.createEvent(INTENT, ctx)).rejects.toThrow('not created: bad start');
});

test('listCalendars maps canEdit/owner/default and mailbox time zone', async () => {
  const { w, calls } = graph([
    ['GET', /^\/me\?/, { status: 200, body: { mail: 'Me@X.com' } }],
    ['GET', /^\/me\/mailboxSettings/, { status: 200, body: { timeZone: 'W. Europe Standard Time' } }],
    ['GET', /^\/me\/calendars\?/, { status: 200, body: { value: [
      { id: 'C1', name: 'Calendar', hexColor: '#0078d4', isDefaultCalendar: true, canEdit: true, owner: { address: 'me@x.com' } },
      { id: 'C2', name: 'Team', hexColor: 'auto', isDefaultCalendar: false, canEdit: false, owner: { address: 'boss@x.com' } },
    ] } }],
  ]);
  const out = await w.listCalendars(ctx);
  expect(out).toEqual([
    { id: 'C1', name: 'Calendar', color: '#0078d4', primary: true, canWrite: true, shared: false, timeZone: 'Europe/Berlin' },
    { id: 'C2', name: 'Team', color: null, primary: false, canWrite: false, shared: true, timeZone: 'Europe/Berlin' },
  ]);
  await w.listCalendars(ctx);
  expect(calls.filter((c) => /mailboxSettings/.test(c.url))).toHaveLength(1);
});

test('document keeps tentative and oof', () => {
  const cal = { id: 'CAL1', name: 'Calendar', color: null };
  const ev = {
    id: 'EV1', start: { dateTime: '2026-10-01T09:00:00.0000000' }, end: { dateTime: '2026-10-01T10:00:00.0000000' },
  };
  const md = (e: object) =>
    (calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: { ...ev, ...e } as never }).metadata as Record<string, unknown>);
  expect(md({ showAs: 'tentative' })).toMatchObject({ status: 'tentative', eventType: 'default', transparency: 'opaque' });
  expect(md({ showAs: 'oof' })).toMatchObject({ status: 'confirmed', eventType: 'outOfOffice' });
  expect(md({ showAs: 'free' })).toMatchObject({ transparency: 'transparent' });
});

test('document reads category from the stamp', () => {
  const cal = { id: 'CAL1', name: 'Calendar', color: null };
  const ev = { id: 'EV1', start: { dateTime: '2026-10-01T09:00:00.0000000' }, end: { dateTime: '2026-10-01T10:00:00.0000000' } };
  const cat = (e: object) =>
    (calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: { ...ev, ...e } as never }).metadata as Record<string, unknown>).category;
  const stamp = { id: 'String {f5d1d8b2-3c4e-4a7b-9d61-6b1c2a9e7f10} Name kiaStamp', value: 'kids|D1' };
  expect(cat({ singleValueExtendedProperties: [stamp], categories: ['Work'] })).toBe('kids');
  expect(cat({ categories: ['Red category', 'PERSONAL'] })).toBe('personal');
  expect(cat({})).toBeNull();
});
