import { GraphClient } from '../graph-client';
import {
  discoverTracked,
  listChildFolders,
  listTopFolders,
  resolveWellKnown,
  type MailFolderNode,
} from '../folders';
import { graphFetch, instantClock, jsonRes, type GraphWorld } from '../testing/harness';

function client(world: GraphWorld) {
  const { fetchFn, calls } = graphFetch(world);
  return {
    c: new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock }),
    calls,
  };
}

const f = (id: string, childFolderCount = 0, parentFolderId = 'ROOT'): MailFolderNode => ({
  id,
  displayName: id.toLowerCase(),
  parentFolderId,
  childFolderCount,
});

//  INBOX ─┬─ A ── A1
//         └─ B
//  ARCH
//  SEARCH (the searchfolders well-known root)
const TREE: GraphWorld['folders'] = {
  top: [[f('INBOX', 2), f('SEARCH', 1)], [f('ARCH')]],
  children: {
    INBOX: [[f('A', 1, 'INBOX')], [f('B', 0, 'INBOX')]],
    A: [f('A1', 0, 'A')],
    B: [],
    ARCH: [],
  },
  wellKnown: {
    searchfolders: f('SEARCH', 1),
    inbox: f('INBOX', 2),
    archive: f('ARCH'),
  },
};

describe('folder discovery', () => {
  it('lists top folders across pages, without the searchfolders root', async () => {
    const { c } = client({ folders: TREE });
    expect((await listTopFolders(c)).map((n) => n.id)).toEqual(['INBOX', 'ARCH']);
  });

  it('tolerates a mailbox with no searchfolders root', async () => {
    const { c } = client({ folders: { ...TREE, wellKnown: {} } });
    expect((await listTopFolders(c)).map((n) => n.id)).toEqual(['INBOX', 'SEARCH', 'ARCH']);
  });

  it('lists child folders across pages', async () => {
    const { c } = client({ folders: TREE });
    expect((await listChildFolders(c, 'INBOX')).map((n) => n.id)).toEqual(['A', 'B']);
  });

  it('discovers a whole subtree, each folder mapped to its root', async () => {
    const { c } = client({ folders: TREE });
    const tracked = await discoverTracked(c, ['INBOX', 'ARCH']);
    expect(Object.fromEntries(tracked)).toEqual({
      INBOX: 'INBOX',
      A: 'INBOX',
      B: 'INBOX',
      A1: 'INBOX',
      ARCH: 'ARCH',
    });
  });

  it('overlapping roots: the first root in order wins', async () => {
    const { c } = client({ folders: TREE });
    const tracked = await discoverTracked(c, ['INBOX', 'A']);
    expect(tracked.get('A')).toBe('INBOX');
    expect(tracked.get('A1')).toBe('INBOX');
  });

  it('does not list children of a folder that has none', async () => {
    const { c, calls } = client({ folders: TREE });
    await discoverTracked(c, ['INBOX']);
    expect(calls.some((u) => u.includes('/B/childFolders'))).toBe(false);
  });

  it('fails closed: a childFolders 500 after retries rejects discovery', async () => {
    const { c } = client({
      folders: TREE,
      custom: (url) => (url.pathname.endsWith('/A/childFolders') ? jsonRes(500, {}) : undefined),
    });
    await expect(discoverTracked(c, ['INBOX'])).rejects.toThrow(/500/);
  });

  it('resolves well-known names; a missing one is omitted', async () => {
    const { c } = client({ folders: TREE });
    const got = await resolveWellKnown(c, ['inbox', 'sentitems', 'archive']);
    expect(Object.keys(got)).toEqual(['inbox', 'archive']);
    expect(got.archive.id).toBe('ARCH');
  });
});
