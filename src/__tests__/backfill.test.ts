/**
 * Backfill suite (enumerate + ingest phases via source.pull()). Adapted
 * from legacy `src/__tests__/ms365-backfill.test.ts`'s
 * enumerateConversations / ingestConversations / runBackfill scenarios onto
 * the v2 `pull()` async generator: one cursor+items commit per yielded
 * Batch instead of legacy's separate sync_state rows.
 */
import { createMs365Source } from '../source';
import { initialDeltaUrl } from '../graph-api';
import type { Batch } from '@kiagent/connector-sdk';
import type { Ms365Cursor } from '../cursor';
import type { Ms365ThreadItem } from '../to-document';
import { collect, graphFetch, graphMsg, instantClock, makeHost, makeSession } from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365ThreadItem>;

const INBOX_START = initialDeltaUrl('INBOX-ID');
const SENT_START = initialDeltaUrl('SENT-ID');

function makeSource(world: Parameters<typeof graphFetch>[0]) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  return { source, calls };
}

describe('backfill: enumerate + ingest', () => {
  it('pages each folder, keeps drafts and every folder (legacy retention), ingests, ends live', async () => {
    const { source, calls } = makeSource({
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
        C3: [graphMsg({ id: 'm3', conversationId: 'C3', subject: 'draft C3' })],
        C4: [graphMsg({ id: 'm5', conversationId: 'C4', subject: 'hello C4' })],
        C5: [graphMsg({ id: 'm4', conversationId: 'C5', parentFolderId: 'JUNK' })],
        C9: [graphMsg({ id: 's1', conversationId: 'C9', subject: 'hello C9' })],
      },
    });
    const { session } = makeSession({ config: { tenantKind: 'personal' } });

    const batches = (await collect(source.pull(session, null))) as B[];

    expect(batches[batches.length - 1]).toEqual({
      phase: 'live',
      items: [],
      cursor: {
        v: 2,
        phase: 'live',
        folders: {
          'INBOX-ID': { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
          'SENT-ID': { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
        },
        pending: [],
        retry: [],
      },
    });

    // Per-folder delta only ever lists that folder's own messages, so the
    // old junk/deleted exclusion never fired on the Inbox/Sent feeds; a
    // JUNK-parented entry here is fixture-only. Drafts now count.
    const allConversationIds = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    expect(allConversationIds.sort()).toEqual(['C1', 'C3', 'C4', 'C5', 'C9']);

    const ingestBatch = batches.find((b) => b.items.length > 0)!;
    expect(ingestBatch.estimateTotal).toBe(5);
    expect(calls.some((u) => u.includes('junkemail'))).toBe(false);
  });

  it('checkpoints a resume cursor after every delta page, and a crash mid-folder resumes from the last page fetched without re-fetching it', async () => {
    // A 3-page inbox feed (p1 -> p2 -> p3-with-deltaLink), sentitems finishes
    // in one page. This proves two things the old per-FOLDER checkpoint
    // (yielded only once, after `walkGraphDelta` fully resolved) could not:
    // (1) a mid-walk cursor appears after page 1, well before the folder's
    //     deltaLink is ever reached, and (2) resuming a crashed run from that
    //     cursor fetches ONLY the remaining pages — page 1 is never re-fetched.
    const world = {
      urls: {
        [INBOX_START]: {
          value: [{ id: 'm1', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false }],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/inbox-p2',
        },
        'https://graph.microsoft.com/v1.0/inbox-p2': {
          value: [{ id: 'm2', conversationId: 'C2', parentFolderId: 'inbox', isDraft: false }],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/inbox-p3',
        },
        'https://graph.microsoft.com/v1.0/inbox-p3': {
          value: [{ id: 'm3', conversationId: 'C3', parentFolderId: 'inbox', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-final',
        },
        [SENT_START]: {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-final',
        },
      },
      conversations: {
        C1: [graphMsg({ id: 'm1', conversationId: 'C1' })],
        C2: [graphMsg({ id: 'm2', conversationId: 'C2' })],
        C3: [graphMsg({ id: 'm3', conversationId: 'C3' })],
      },
    };

    // --- First run: abandon it ("crash") right after page 1's checkpoint ---
    const { source: firstRun, calls: firstCalls } = makeSource(world);
    const { session: firstSession } = makeSession({ config: { tenantKind: 'personal' } });
    const gen = firstRun.pull(firstSession, null)[Symbol.asyncIterator]();

    const firstBatch = (await gen.next()).value as B;
    // The very first yielded batch is already a mid-folder checkpoint: the
    // inbox delta walk has only consumed page 1 (deltaLink not yet reached),
    // proving checkpoints are produced per-page, not per-folder.
    expect(firstBatch.phase).toBe('backfill');
    expect(firstBatch.items).toEqual([]);
    expect(firstBatch.cursor).toEqual({
      v: 2,
      phase: 'enumerate',
      folders: {
        'INBOX-ID': { next: 'https://graph.microsoft.com/v1.0/inbox-p2' },
        'SENT-ID': { next: SENT_START },
      },
      pending: ['C1'],
      retry: [],
    });
    // Discovery ran first; of the delta pages, only page 1 was fetched.
    expect(firstCalls.filter((u) => u.includes('/messages/delta') || u.includes('inbox-p'))).toEqual([
      INBOX_START,
    ]);
    // Abandon the generator here — this is the simulated crash. Page 2 and 3
    // are never fetched by this run.
    await gen.return?.();

    // --- Second run: fresh source/session, resumed from the captured cursor ---
    const { source: secondRun, calls: secondCalls } = makeSource(world);
    const { session: secondSession } = makeSession({ config: { tenantKind: 'personal' } });
    const batches = (await collect(
      secondRun.pull(secondSession, firstBatch.cursor),
    )) as B[];

    // Page 1's URL must never be re-fetched by the resumed run.
    expect(secondCalls).not.toContain(INBOX_START);
    expect(secondCalls).toContain('https://graph.microsoft.com/v1.0/inbox-p2');
    expect(secondCalls).toContain('https://graph.microsoft.com/v1.0/inbox-p3');

    const allIds = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    // C1 (from the abandoned first run's page) is still pending and gets
    // ingested; C2 and C3 come from the resumed pages.
    expect(allIds.sort()).toEqual(['C1', 'C2', 'C3']);
  });

  it('resumes from a saved enumerate cursor', async () => {
    const { source } = makeSource({
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
      v: 2,
      phase: 'enumerate',
      folders: {
        'INBOX-ID': { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        'SENT-ID': { next: 'https://graph.microsoft.com/v1.0/sent-resume' },
      },
      pending: ['C0'],
      retry: [],
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
      v: 2,
      phase: 'ingest',
      folders: {
        'INBOX-ID': { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        'SENT-ID': { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
      },
      pending: ['C1'],
      total: 5,
      retry: [],
    };

    const batches = (await collect(source.pull(session, resumeCursor))) as B[];
    expect(batches.some((b) => b.items.some((i) => i.conversationId === 'C1'))).toBe(true);
    expect(calls.some((u) => u.includes('/messages/delta'))).toBe(false); // no re-enumeration
    const last = batches[batches.length - 1];
    expect(last.cursor).toEqual({
      v: 2,
      phase: 'live',
      folders: {
        'INBOX-ID': { delta: 'https://graph.microsoft.com/v1.0/inbox-final' },
        'SENT-ID': { delta: 'https://graph.microsoft.com/v1.0/sent-final' },
      },
      pending: [],
      retry: [],
    });
  });

  it('continues past a per-conversation failure, keeping it in retry', async () => {
    const { source } = makeSource({
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
      v: 2,
      phase: 'ingest',
      folders: { 'INBOX-ID': { delta: 'x' }, 'SENT-ID': { delta: 'y' } },
      pending: ['BAD', 'GOOD'],
      total: 2,
      retry: [],
    };

    const batches = (await collect(source.pull(session, resumeCursor))) as B[];
    const ids = batches.flatMap((b) => b.items.map((i) => i.conversationId));
    expect(ids).toEqual(['GOOD']);
    // The failure is kept for the next pull, not logged and lost.
    expect(batches[batches.length - 1].cursor.retry).toEqual([{ id: 'BAD', n: 1 }]);
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
      v: 2,
      phase: 'ingest',
      folders: { 'INBOX-ID': { delta: 'x' }, 'SENT-ID': { delta: 'y' } },
      pending: ['DEAD1', 'DEAD2', 'NEVER'],
      total: 3,
      retry: [],
    };

    await expect(collect(source.pull(session, resumeCursor))).rejects.toThrow(/401/);
    // 'NEVER' sits in the batch AFTER the one whose auth failure propagated
    // (INGEST_CONCURRENCY=2) — it must never have been fetched.
    expect(neverCalled).toBe(false);
  });
});
