import type {
  FolderNode,
  FolderPickerSpec,
  FolderScopeUpdate,
  FolderSelectionChannel,
} from '@kiagent/connector-sdk';
import { createMs365Source } from '../source';
import type { Ms365Cursor } from '../cursor';
import {
  graphFetch,
  instantClock,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

const node = (id: string, displayName: string, parentFolderId = 'ROOT', childFolderCount = 0) => ({
  id,
  displayName,
  parentFolderId,
  childFolderCount,
});

//  ROOT ─┬─ INBOX-ID ── PROJ ── PROJ-A
//        ├─ SENT-ID
//        ├─ ARCHIVE-ID
//        ├─ JUNK-ID
//        └─ SEARCH-ID (searchfolders)
const TREE: GraphWorld['folders'] = {
  top: [
    node('INBOX-ID', 'Inbox', 'ROOT', 1),
    node('SENT-ID', 'Sent Items'),
    node('ARCHIVE-ID', 'Archive'),
    node('JUNK-ID', 'Junk Email'),
    node('SEARCH-ID', 'Search Folders'),
  ],
  children: {
    'INBOX-ID': [node('PROJ', 'Projects', 'INBOX-ID', 1)],
    PROJ: [node('PROJ-A', 'Alpha', 'PROJ')],
    'SENT-ID': [],
    'ARCHIVE-ID': [],
    'JUNK-ID': [],
  },
  wellKnown: {
    inbox: node('INBOX-ID', 'Inbox', 'ROOT', 1),
    sentitems: node('SENT-ID', 'Sent Items'),
    archive: node('ARCHIVE-ID', 'Archive'),
    junkemail: node('JUNK-ID', 'Junk Email'),
    searchfolders: node('SEARCH-ID', 'Search Folders'),
    msgfolderroot: node('ROOT', 'Top of Information Store', ''),
    'PROJ-A': node('PROJ-A', 'Alpha', 'PROJ'),
    PROJ: node('PROJ', 'Projects', 'INBOX-ID', 1),
    'INBOX-ID': node('INBOX-ID', 'Inbox', 'ROOT', 1),
    'SENT-ID': node('SENT-ID', 'Sent Items'),
    'ARCHIVE-ID': node('ARCHIVE-ID', 'Archive'),
    'JUNK-ID': node('JUNK-ID', 'Junk Email'),
  },
};

const FOLDER_MESSAGES = {
  'INBOX-ID': ['C-both', 'C-inbox'],
  PROJ: ['C-proj'],
  'PROJ-A': [],
  'SENT-ID': ['C-both', 'C-sent'],
  'ARCHIVE-ID': ['C-arch'],
};

const INBOX = { id: 'INBOX-ID', name: 'Inbox' };
const SENT = { id: 'SENT-ID', name: 'Sent Items' };
const ARCHIVE = { id: 'ARCHIVE-ID', name: 'Archive' };

async function manage(opts: {
  pick: string[];
  folderRoots?: Array<{ id: string; name: string }>;
  calendarRoots?: Array<{ id: string; name: string }>;
  calendars?: unknown;
  calendarStatus?: number;
  cursor?: unknown;
}) {
  const { fetchFn, calls } = graphFetch({
    folders: TREE,
    folderMessages: FOLDER_MESSAGES,
    calendars: opts.calendars,
    calendarStatus: opts.calendarStatus,
  });
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  const { session } = makeSession({
    config: {
      tenantKind: 'work',
      ...(opts.folderRoots ? { folderRoots: opts.folderRoots } : {}),
      ...(opts.calendarRoots ? { calendarRoots: opts.calendarRoots } : {}),
    },
    cursor: opts.cursor,
  });
  const seen: { spec?: FolderPickerSpec; roots?: FolderNode[]; calendars?: FolderNode[] } = {};
  const channel: FolderSelectionChannel = {
    status: () => {},
    pickFolders: async (spec) => {
      seen.spec = spec;
      seen.roots = await spec.roots(spec.modes[0].key);
      seen.calendars = spec.modes[1] ? await spec.roots(spec.modes[1].key) : [];
      const known = new Map<string, FolderNode>();
      for (const n of [
        ...seen.roots,
        ...seen.calendars,
        ...(await spec.children('INBOX-ID')),
        ...(await spec.children('PROJ')),
      ])
        known.set(n.id, n);
      return opts.pick.map((id) => known.get(id)!);
    },
  };
  const run = () => source.manageFolders!(session, channel) as Promise<FolderScopeUpdate<Ms365Cursor>>;
  return { run, calls, seen };
}

const refIds = (u: FolderScopeUpdate<Ms365Cursor>) => (u.archiveRefs ?? []).map((r) => r.externalId).sort();
const listings = (calls: string[]) => calls.filter((c) => c.includes('$select=conversationId'));

describe('manageFolders', () => {
  it('the picker: Mail folders mode, the note, Junk flagged, search folders hidden, selection + ancestors expanded', async () => {
    const m = await manage({ pick: ['INBOX-ID'], folderRoots: [INBOX, { id: 'PROJ-A', name: 'Alpha' }] });
    await m.run();
    const spec = m.seen.spec!;
    expect(spec.modes).toEqual([
      { key: 'mail', label: 'Mail folders' },
      { key: 'calendars', label: 'Calendars' },
    ]);
    expect(spec.multiSelect).toBe(true);
    expect(spec.purpose).toBe('manage');
    expect(spec.note).toBe('Mail outside the selected folders will be removed from the index');
    expect(spec.count).toBeUndefined();
    expect(m.seen.roots!.map((n) => n.name)).toEqual([
      'Inbox',
      'Sent Items',
      'Archive',
      'Junk Email (may contain phishing)',
    ]);
    expect(m.seen.roots![0].hasChildren).toBe(true);
    expect(spec.selected!.map((n) => n.id)).toEqual(['INBOX-ID', 'PROJ-A']);
    // Ancestors of the selected roots, never a selected root itself.
    expect([...(spec.expand ?? [])].sort()).toEqual(['PROJ']);
  });

  it('removing Inbox archives Inbox-only threads, never one that also lives in Sent', async () => {
    const m = await manage({ pick: ['SENT-ID'], folderRoots: [INBOX, SENT] });
    const up = await m.run();
    expect(refIds(up)).toEqual(['C-inbox', 'C-proj']);
    expect(up.archiveScopeRootIds).toEqual([]);
    expect(up.reattributeScopeRoots).toEqual([]);
    expect(up.config.folderRoots).toEqual([SENT]);
  });

  it('pure widening lists nothing and archives nothing; order = retained, then new', async () => {
    const m = await manage({ pick: ['ARCHIVE-ID', 'INBOX-ID'], folderRoots: [INBOX] });
    const up = await m.run();
    expect(listings(m.calls)).toEqual([]);
    expect(up.archiveRefs).toEqual([]);
    expect(up.config.folderRoots).toEqual([INBOX, ARCHIVE]);
  });

  it("a legacy account's first Save persists folderRoots and archives nothing (core's first-declaration allowance reconciles)", async () => {
    const m = await manage({ pick: ['INBOX-ID'] });
    const up = await m.run();
    expect(listings(m.calls)).toEqual([]);
    expect(up.archiveRefs).toEqual([]);
    expect(up.config).toEqual({ tenantKind: 'work', folderRoots: [INBOX] });
    // The picker pre-selected the legacy default, Inbox + Sent Items.
    expect(m.seen.spec!.selected!.map((n) => n.id)).toEqual(['INBOX-ID', 'SENT-ID']);
  });

  it('removed folder states leave the cursor; kept and newly tracked ones stay / start', async () => {
    const cursor: Ms365Cursor = {
      v: 2,
      attachments: 1,
      phase: 'live',
      folders: {
        'INBOX-ID': { delta: 'DI' },
        PROJ: { delta: 'DP' },
        'PROJ-A': { delta: 'DA' },
        'SENT-ID': { delta: 'DS' },
      },
      pending: ['x'],
      retry: [],
    };
    const up = await (await manage({ pick: ['SENT-ID', 'ARCHIVE-ID'], folderRoots: [INBOX, SENT], cursor })).run();
    expect(Object.keys(up.cursor!.folders)).toEqual(['SENT-ID', 'ARCHIVE-ID']);
    expect(up.cursor!.folders['SENT-ID']).toEqual({ delta: 'DS' });
    expect(up.cursor!.pending).toEqual(['x']);
  });

  it('a cursor from before attachments is dropped (the next pull re-enumerates)', async () => {
    const up = await (
      await manage({
        pick: ['INBOX-ID', 'SENT-ID'],
        cursor: { phase: 'live', folders: { inbox: { delta: 'DI' }, sentitems: { delta: 'DS' } } },
      })
    ).run();
    expect(up.cursor).toBeNull();
  });

  it('a null cursor stays null', async () => {
    const up = await (await manage({ pick: ['INBOX-ID'], folderRoots: [INBOX] })).run();
    expect(up.cursor).toBeNull();
  });

  it('an empty pick rejects', async () => {
    await expect((await manage({ pick: [], folderRoots: [INBOX] })).run()).rejects.toThrow(
      'ms365: no mail folders selected',
    );
  });

  it('the stored name drops the Junk warning suffix', async () => {
    const up = await (await manage({ pick: ['INBOX-ID', 'JUNK-ID'], folderRoots: [INBOX] })).run();
    expect(up.config.folderRoots).toEqual([INBOX, { id: 'JUNK-ID', name: 'Junk Email' }]);
  });
});

describe('manageFolders: the Calendars tab (spec 2026-09-30 §6)', () => {
  const CALS = {
    value: [
      { id: 'CAL', name: 'Calendar', isDefaultCalendar: true, canEdit: true },
      { id: 'TEAM', name: 'Team', canEdit: true },
      { id: 'BDAY', name: 'Birthdays', canEdit: false },
    ],
  };
  const cursorWithCal = {
    v: 2,
    attachments: 1,
    phase: 'live',
    folders: {},
    pending: [],
    retry: [],
    calendar: { since: '2025-10-01T00:00:00.000Z', cals: { CAL: { e1: 'x' }, TEAM: { e2: 'y' } } },
  };

  it('lists the calendars flat; the default and owned ones are pre-checked', async () => {
    const m = await manage({ pick: ['INBOX-ID', 'CAL'], folderRoots: [INBOX], calendars: CALS });
    await m.run();
    expect(m.seen.calendars).toEqual([
      { id: 'CAL', name: 'Calendar', hasChildren: false },
      { id: 'TEAM', name: 'Team', hasChildren: false },
      { id: 'BDAY', name: 'Birthdays', hasChildren: false },
    ]);
    expect(m.seen.spec!.selected!.map((n) => n.id)).toEqual(['INBOX-ID', 'CAL', 'TEAM']);
  });

  it('unticking every calendar turns the calendar off: [] roots, their events archived, cursor pruned', async () => {
    const m = await manage({ pick: ['INBOX-ID'], folderRoots: [INBOX], calendars: CALS, cursor: cursorWithCal });
    const up = await m.run();
    expect(up.config.calendarRoots).toEqual([]);
    expect(up.config.folderRoots).toEqual([INBOX]);
    expect(up.archiveScopeRootIds).toEqual(['CAL', 'TEAM']);
    expect(up.cursor!.calendar!.cals).toEqual({});
  });

  it('no Calendars.Read consent yet (403): a mail Save leaves the calendar setting alone', async () => {
    const m = await manage({ pick: ['INBOX-ID'], folderRoots: [INBOX], calendarStatus: 403 });
    const up = await m.run();
    expect(m.seen.calendars).toEqual([]);
    expect(up.config).not.toHaveProperty('calendarRoots');
    expect(up.archiveScopeRootIds).toEqual([]);
  });

  it('keeping a calendar but no mail folder is refused', async () => {
    const m = await manage({ pick: ['CAL'], folderRoots: [INBOX], calendars: CALS });
    await expect(m.run()).rejects.toThrow('ms365: no mail folders selected');
  });

  it('adding a calendar tracks it and archives nothing', async () => {
    const m = await manage({
      pick: ['INBOX-ID', 'CAL', 'BDAY'],
      folderRoots: [INBOX],
      calendarRoots: [{ id: 'CAL', name: 'Calendar' }],
      calendars: CALS,
    });
    const up = await m.run();
    expect(up.config.calendarRoots).toEqual([
      { id: 'CAL', name: 'Calendar' },
      { id: 'BDAY', name: 'Birthdays' },
    ]);
    expect(up.archiveScopeRootIds).toEqual([]);
  });
});
