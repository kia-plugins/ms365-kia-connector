import { GraphClient } from '../graph-client';
import { configuredRoots, effectiveRoots } from '../scope';
import { graphFetch, instantClock } from '../testing/harness';

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
