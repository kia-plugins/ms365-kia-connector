import { rescope, type Ms365Cursor } from '../cursor';


describe('rescope', () => {
  it('drops untracked folder states and starts new tracked folders from their initial delta', () => {
    const cur: Ms365Cursor = {
      v: 2,
      phase: 'live',
      folders: { KEEP: { delta: 'DK' }, GONE: { delta: 'DG' } },
      pending: ['c1'],
      retry: [{ id: 'c2', n: 1 }],
    };
    const out = rescope(cur, new Map([['KEEP', 'KEEP'], ['NEW/1', 'KEEP']]));
    expect(out.folders).toEqual({
      KEEP: { delta: 'DK' },
      'NEW/1': {
        next: 'https://graph.microsoft.com/v1.0/me/mailFolders/NEW%2F1/messages/delta?$select=id,conversationId,parentFolderId,isDraft&$top=100',
      },
    });
    expect(out.pending).toEqual(['c1']);
    expect(out.retry).toEqual([{ id: 'c2', n: 1 }]);
    expect(out.phase).toBe('live');
  });
});
