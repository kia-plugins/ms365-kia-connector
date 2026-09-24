import { GraphClient } from '../graph-client';
import { configuredRoots, effectiveRoots, resolveScope } from '../scope';
import { graphFetch, instantClock, jsonRes } from '../testing/harness';

const client = () =>
  new GraphClient({ fetch: graphFetch().fetchFn, getToken: async () => 'tok', ...instantClock });

describe('scope config', () => {
  it('a config without a folderRoots array is legacy (null)', () => {
    expect(configuredRoots({})).toBeNull();
    expect(configuredRoots({ tenantKind: 'work', folderRoots: 'x' })).toBeNull();
  });

  it('reads folderRoots, dropping malformed entries', () => {
    expect(
      configuredRoots({ folderRoots: [{ id: 'F1', name: 'Projects' }, { id: 7 }, null] }),
    ).toEqual([{ id: 'F1', name: 'Projects' }]);
  });

  it('legacy accounts enumerate the well-known Inbox + Sent Items', async () => {
    expect(await effectiveRoots(client(), {})).toEqual({
      roots: [
        { id: 'INBOX-ID', name: 'Inbox' },
        { id: 'SENT-ID', name: 'Sent Items' },
      ],
      legacy: true,
    });
  });

  it('a configured selection is used as-is, without a Graph call', async () => {
    const roots = [{ id: 'F1', name: 'Projects' }];
    expect(await effectiveRoots(client(), { folderRoots: roots })).toEqual({
      roots,
      legacy: false,
    });
  });
});

describe('resolveScope', () => {
  it('roots + legacy flag + the discovered tracked map, in one call', async () => {
    const c = new GraphClient({
      fetch: graphFetch({
        folders: {
          top: [],
          children: { F1: [{ id: 'F2', displayName: 'Sub', childFolderCount: 0 }] },
        },
      }).fetchFn,
      getToken: async () => 'tok',
      ...instantClock,
    });
    const scope = await resolveScope(c, { folderRoots: [{ id: 'F1', name: 'Projects' }] });
    expect(scope.legacy).toBe(false);
    expect(scope.roots).toEqual([{ id: 'F1', name: 'Projects' }]);
    expect(Object.fromEntries(scope.tracked)).toEqual({ F1: 'F1', F2: 'F1' });
  });
});

describe('resolveScope fails closed', () => {
  it('when not one selected folder exists any more, it throws instead of tracking nothing', async () => {
    const c = new GraphClient({
      fetch: graphFetch({
        custom: (url) =>
          url.pathname.endsWith('/childFolders')
            ? jsonRes(404, { error: { code: 'ErrorItemNotFound' } })
            : undefined,
      }).fetchFn,
      getToken: async () => 'tok',
      ...instantClock,
    });
    await expect(
      resolveScope(c, { folderRoots: [{ id: 'F1', name: 'Projects' }, { id: 'F2', name: 'Other' }] }),
    ).rejects.toThrow('ms365: none of the selected mail folders exists');
  });
});
