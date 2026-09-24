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

  it('a leaf folder deleted upstream mid-pass is skipped, not fatal', async () => {
    const { run, logs } = reconcile({
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
    expect(ids(await run())).toEqual(['C-both', 'C-inbox', 'C-inbox2']);
    expect(logs.some((l) => l.level === 'warn' && l.msg.includes('LEAF'))).toBe(true);
  });

  it('refuses to run for an account with no declared scope', async () => {
    await expect(reconcile(WORLD, null).run()).rejects.toThrow(
      'ms365: reconcile without declared scope',
    );
  });
});
