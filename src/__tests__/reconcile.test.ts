import type { ExternalRef } from '@kiagent/connector-sdk';
import { createMs365Source } from '../source';
import {
  collect,
  DEFAULT_FOLDERS,
  graphFetch,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

const ROOTS = [
  { id: 'INBOX-ID', name: 'Inbox' },
  { id: 'SENT-ID', name: 'Sent Items' },
];

function reconcile(world: GraphWorld, folderRoots: unknown = ROOTS) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  const { session, logs } = makeSession({
    config: folderRoots === null ? {} : { folderRoots },
  });
  return { calls, logs, run: () => collect(source.reconcile!(session)) };
}

const ids = (pages: ExternalRef[][]) => [...new Set(pages.flat().map((r) => r.externalId))].sort();

const WORLD: GraphWorld = {
  folders: { ...DEFAULT_FOLDERS, children: { ...DEFAULT_FOLDERS.children, 'DELETED-ID': [] } },
  folderMessages: {
    'INBOX-ID': [['C-both', 'C-inbox'], ['C-inbox2']],
    'SENT-ID': ['C-both'],
    'DELETED-ID': ['C-trash'],
  },
};

describe('reconcile', () => {
  it('lists every conversation with a message in the tracked folders, across pages', async () => {
    const { run, logs } = reconcile(WORLD);
    const pages = await run();
    expect(ids(pages)).toEqual(['C-both', 'C-inbox', 'C-inbox2']);
    expect(pages.flat().every((r) => r.type === 'email.thread')).toBe(true);
    expect(logs.some((l) => l.level === 'info' && /requests/.test(l.msg))).toBe(true);
  });

  it('a conversation only in untracked Deleted Items is not listed…', async () => {
    expect(ids(await reconcile(WORLD).run())).not.toContain('C-trash');
  });

  it('…and is listed once Deleted Items is tracked', async () => {
    const { run } = reconcile(WORLD, [...ROOTS, { id: 'DELETED-ID', name: 'Deleted Items' }]);
    expect(ids(await run())).toContain('C-trash');
  });

  it('a discovery failure rejects before anything is yielded', async () => {
    const { run } = reconcile({
      ...WORLD,
      custom: (url) => (url.pathname.endsWith('/INBOX-ID/childFolders') ? jsonRes(500, {}) : undefined),
    });
    await expect(run()).rejects.toThrow(/500/);
  });

  it('a folder that 404s between discovery and listing fails the pass — never read as empty', async () => {
    const { run } = reconcile({
      folders: {
        ...DEFAULT_FOLDERS,
        children: {
          ...DEFAULT_FOLDERS.children,
          'INBOX-ID': [{ id: 'LEAF', displayName: 'Leaf', childFolderCount: 0 }],
        },
      },
      folderMessages: WORLD.folderMessages,
      custom: (url) =>
        url.pathname.endsWith('/LEAF/messages') ? jsonRes(404, { error: { code: 'ErrorItemNotFound' } }) : undefined,
    });
    await expect(run()).rejects.toThrow(/ErrorItemNotFound/);
  });

  it('a mailbox-level 404 on a folder listing rejects the pass', async () => {
    const { run } = reconcile({
      ...WORLD,
      custom: (url) =>
        url.pathname.endsWith('/SENT-ID/messages')
          ? jsonRes(404, { error: { code: 'MailboxNotEnabledForRESTAPI' } })
          : undefined,
    });
    await expect(run()).rejects.toThrow(/MailboxNotEnabledForRESTAPI/);
  });

  it('refuses to run for an account with no declared scope', async () => {
    await expect(reconcile(WORLD, null).run()).rejects.toThrow(
      'ms365: reconcile without declared scope',
    );
  });
});

describe('reconcile: calendar events (spec 2026-09-30 §6)', () => {
  const DAY = 86_400_000;
  const since = new Date(Date.now() - 365 * DAY).toISOString();
  const old = new Date(Date.now() - 200 * DAY).toISOString().replace('Z', '');
  const CALS = { value: [{ id: 'C1', name: 'Calendar', isDefaultCalendar: true }] };

  function withCursor(world: GraphWorld, calendar?: unknown) {
    const { fetchFn, calls } = graphFetch(world);
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { session } = makeSession({
      config: { folderRoots: ROOTS },
      cursor: { v: 2, attachments: 1, phase: 'live', folders: {}, pending: [], retry: [], calendar },
    });
    return { calls, run: () => collect(source.reconcile!(session)) };
  }

  it('lists every calendar event from the fixed first-pull start, so an old one is kept', async () => {
    const { run, calls } = withCursor(
      {
        ...WORLD,
        calendars: CALS,
        calendarViews: { C1: [{ id: 'OLD', start: { dateTime: old }, end: { dateTime: old } }] },
      },
      { since, cals: { C1: {} } },
    );
    const refs = (await run()).flat();
    expect(refs).toContainEqual({ externalId: 'C1:OLD', type: 'calendar.event' });
    expect(refs.some((r) => r.type === 'email.thread')).toBe(true);
    const view = calls.find((u) => u.includes('/calendarView'))!;
    expect(new URL(view).searchParams.get('startDateTime')).toBe(since);
  });

  it('no consent and never synced (403, no calendar cursor): mail refs only', async () => {
    const refs = (await withCursor({ ...WORLD, calendarStatus: 403 }).run()).flat();
    expect(refs.every((r) => r.type === 'email.thread')).toBe(true);
    expect(refs.length).toBeGreaterThan(0);
  });

  it('a 403 after calendars synced fails the pass (nothing is archived)', async () => {
    await expect(
      withCursor({ ...WORLD, calendarStatus: 403 }, { since, cals: { C1: {} } }).run(),
    ).rejects.toThrow();
  });

  it('a 500 fails the pass', async () => {
    await expect(withCursor({ ...WORLD, calendarStatus: 500 }).run()).rejects.toThrow();
  });
});
