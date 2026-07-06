/**
 * Backfill suite (enumerate + ingest phases via source.pull()). Adapted
 * from legacy `src/__tests__/ms365-backfill.test.ts`'s
 * enumerateConversations / ingestConversations / runBackfill scenarios onto
 * the v2 `pull()` async generator: one cursor+items commit per yielded
 * Batch instead of legacy's separate sync_state rows.
 */
import { createMs365Source } from '../source';
import { initialDeltaUrl } from '../graph-api';
import type { Batch } from '../kiagent-contracts';
import type { Ms365Cursor } from '../cursor';
import type { Ms365ThreadItem } from '../to-document';
import { collect, graphFetch, graphMsg, instantClock, makeHost, makeSession } from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365ThreadItem>;

const INBOX_START = initialDeltaUrl('inbox');
const SENT_START = initialDeltaUrl('sentitems');

function makeSource(world: Parameters<typeof graphFetch>[0]) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  return { source, calls };
}

describe('backfill: enumerate + ingest', () => {
  it('pages each folder, skips drafts + excluded folders, ingests, ends live', async () => {
    const { source, calls } = makeSource({
      junkFolderId: 'JUNK',
      trashFolderId: 'TRASH',
      urls: {
        [INBOX_START]: {
          value: [
            { id: 'm1', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false },
            { id: 'm2', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false },
            { id: 'm3', conversationId: 'C3', parentFolderId: 'inbox', isDraft: true },
            { id: 'm4', conversationId: 'C5', parentFolderId: 'JUNK', isDraft: false },
          ],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/inbox-p2',
        },
        'https://graph.microsoft.com/v1.0/inbox-p2': {
          value: [{ id: 'm5', conversationId: 'C4', parentFolderId: 'inbox', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-final',
        },
        [SENT_START]: {
          value: [{ id: 's1', conversationId: 'C9', parentFolderId: 'sentitems', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-final',
        },
      },
      conversations: {
        C1: [graphMsg({ id: 'm1', conversationId: 'C1', subject: 'hello C1' })],
        C4: [graphMsg({ id: 'm5', conversationId: 'C4', subject: 'hello C4' })],
        C9: [graphMsg({ id: 's1', conversationId: 'C9', subject: 'hello C9' })],
      },
    });
    const { session } = makeSession({ config: { tenantKind: 'personal' } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[batches.length - 1]).toEqual({
      phase: 'live',
      items: [],
      cursor: {
        phase: 'live',
        folders: {
          inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
          sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
        },
      },
    });

    const allConversationIds = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    expect(allConversationIds.sort()).toEqual(['C1', 'C4', 'C9']);

    const ingestBatch = batches.find((b) => b.items.length > 0)!;
    expect(ingestBatch.estimateTotal).toBe(3);
    expect(calls.some((u) => u.includes('junkemail'))).toBe(true);
    expect(calls.some((u) => u.includes('deleteditems'))).toBe(true);
  });

  it('resumes from a saved enumerate cursor', async () => {
    const { source } = makeSource({
      junkFolderId: 'JUNK',
      trashFolderId: 'TRASH',
      urls: {
        'https://graph.microsoft.com/v1.0/sent-resume': {
          value: [{ id: 's9', conversationId: 'C9', parentFolderId: 'sentitems', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-final',
        },
      },
      conversations: { C0: [graphMsg({ conversationId: 'C0' })], C9: [graphMsg({ conversationId: 'C9' })] },
    });
    const { session } = makeSession();
    const resumeCursor: Ms365Cursor = {
      phase: 'enumerate',
      folders: {
        inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        sentitems: { next: 'https://graph.microsoft.com/v1.0/sent-resume' },
      },
      pending: ['C0'],
    };

    const batches = (await collect(source.pull(session, resumeCursor))) as B[];
    const allIds = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    expect(allIds.sort()).toEqual(['C0', 'C9']);
  });

  it('resumes from a saved ingest cursor without re-enumerating', async () => {
    const { source, calls } = makeSource({
      conversations: { C1: [graphMsg({ conversationId: 'C1' })] },
    });
    const { session } = makeSession();
    const resumeCursor: Ms365Cursor = {
      phase: 'ingest',
      folders: {
        inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
      },
      pending: ['C1'],
      total: 5,
    };

    const batches = (await collect(source.pull(session, resumeCursor))) as B[];
    expect(batches.some((b) => b.items.some((i) => i.conversationId === 'C1'))).toBe(true);
    expect(calls.some((u) => u.includes('junkemail'))).toBe(false); // no re-enumeration
    const last = batches[batches.length - 1];
    expect(last.cursor).toEqual({
      phase: 'live',
      folders: {
        inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
      },
    });
  });

  it('continues past a per-conversation failure', async () => {
    const { source } = makeSource({
      junkFolderId: 'JUNK',
      trashFolderId: 'TRASH',
      urls: {
        [INBOX_START]: { value: [], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-final' },
        [SENT_START]: { value: [], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-final' },
      },
      conversations: { GOOD: [graphMsg({ conversationId: 'GOOD' })] },
      custom: (url) => {
        if (
          url.pathname === '/v1.0/me/messages' &&
          (url.searchParams.get('$filter') ?? '').includes("'BAD'")
        ) {
          return { status: 500, statusText: '', headers: {}, body: new TextEncoder().encode('oops') };
        }
        return undefined;
      },
    });
    const { session } = makeSession();
    const resumeCursor: Ms365Cursor = {
      phase: 'ingest',
      folders: {
        inbox: { delta: 'x' },
        sentitems: { delta: 'y' },
      },
      pending: ['BAD', 'GOOD'],
      total: 2,
    };

    const batches = (await collect(source.pull(session, resumeCursor))) as B[];
    const ids = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    expect(ids).toEqual(['GOOD']);
  }, 30_000);

  it('rethrows an auth error instead of churning through the remaining conversations', async () => {
    let neverCalled = false;
    const { source } = makeSource({
      custom: (url) => {
        const filter = url.searchParams.get('$filter') ?? '';
        if (filter.includes("'NEVER'")) neverCalled = true;
        if (filter.includes("'DEAD1'") || filter.includes("'DEAD2'") || filter.includes("'NEVER'")) {
          return {
            status: 401,
            statusText: '',
            headers: {},
            body: new TextEncoder().encode('InvalidAuthenticationToken'),
          };
        }
        return undefined;
      },
    });
    const { session } = makeSession();
    const resumeCursor: Ms365Cursor = {
      phase: 'ingest',
      folders: { inbox: { delta: 'x' }, sentitems: { delta: 'y' } },
      pending: ['DEAD1', 'DEAD2', 'NEVER'],
      total: 3,
    };

    await expect(collect(source.pull(session, resumeCursor))).rejects.toThrow(/401/);
    // 'NEVER' sits in the batch AFTER the one whose auth failure propagated
    // (INGEST_CONCURRENCY=2) — it must never have been fetched.
    expect(neverCalled).toBe(false);
  });
});
