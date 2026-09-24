import { migrateCursor, rescope, type Ms365Cursor } from '../cursor';

const IDS = { inbox: 'INBOX-ID', sentitems: 'SENT-ID' };

describe('cursor v2 migration (lossless)', () => {
  it('enumerate: legacy folder keys become ids, links and pending kept', () => {
    expect(
      migrateCursor(
        {
          phase: 'enumerate',
          folders: { inbox: { next: 'N1' }, sentitems: { delta: 'D2' } },
          pending: ['c1'],
        },
        IDS,
      ),
    ).toEqual({
      v: 2,
      phase: 'enumerate',
      folders: { 'INBOX-ID': { next: 'N1' }, 'SENT-ID': { delta: 'D2' } },
      pending: ['c1'],
      retry: [],
    });
  });

  it('ingest keeps pending and total', () => {
    expect(
      migrateCursor(
        {
          phase: 'ingest',
          folders: { inbox: { delta: 'D1' }, sentitems: { delta: 'D2' } },
          pending: ['c1', 'c2'],
          total: 9,
        },
        IDS,
      ),
    ).toEqual({
      v: 2,
      phase: 'ingest',
      folders: { 'INBOX-ID': { delta: 'D1' }, 'SENT-ID': { delta: 'D2' } },
      pending: ['c1', 'c2'],
      total: 9,
      retry: [],
    });
  });

  it('live keeps the delta links', () => {
    expect(
      migrateCursor({ phase: 'live', folders: { inbox: { delta: 'D1' }, sentitems: { delta: 'D2' } } }, IDS),
    ).toEqual({
      v: 2,
      phase: 'live',
      folders: { 'INBOX-ID': { delta: 'D1' }, 'SENT-ID': { delta: 'D2' } },
      pending: [],
      retry: [],
    });
  });

  it('v2 passes through; null stays null', () => {
    const v2: Ms365Cursor = { v: 2, phase: 'live', folders: {}, pending: [], retry: [] };
    expect(migrateCursor(v2, IDS)).toBe(v2);
    expect(migrateCursor(null, IDS)).toBeNull();
  });
});

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
