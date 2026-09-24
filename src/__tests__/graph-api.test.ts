/**
 * Ported from legacy `src/__tests__/ms365-client.test.ts`
 * (resolveExcludedFolderIds, fetchConversationMessages paging/sort) and
 * `src/__tests__/ms-shared-walk-delta.test.ts` (walkGraphDelta), reshaped
 * onto GraphClient / the fake-graph harness.
 */
import { GraphClient } from '../graph-client';
import {
  accumulate,
  fetchConversationMessages,
  walkGraphDelta,
  type GraphDeltaPage,
  type Ms365DeltaMessage,
} from '../graph-api';
import { graphFetch, instantClock } from '../testing/harness';

function client(fetchFn: ReturnType<typeof graphFetch>['fetchFn']) {
  return new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
}

describe('fetchConversationMessages', () => {
  it('pages messages for a single conversationId and sorts client-side, oldest first', async () => {
    const { fetchFn } = graphFetch({
      conversations: {
        C1: [
          [{ id: 'm2', conversationId: 'C1', subject: 'b', receivedDateTime: '2026-05-20T11:00:00Z' }],
          [{ id: 'm1', conversationId: 'C1', subject: 'a', receivedDateTime: '2026-05-20T10:00:00Z' }],
        ],
      },
    });
    const msgs = await fetchConversationMessages(client(fetchFn), 'C1');
    expect(msgs.map((m) => m.id)).toEqual(['m1', 'm2']);
  });

  it('returns empty array for a fixture with zero messages (deleted conversation)', async () => {
    const { fetchFn } = graphFetch({ conversations: { GONE: [] } });
    const msgs = await fetchConversationMessages(client(fetchFn), 'GONE');
    expect(msgs).toEqual([]);
  });

  it('never sends $orderby (Graph rejects it combined with a conversationId filter)', async () => {
    const { fetchFn, calls } = graphFetch({
      conversations: { C1: [{ id: 'm1', conversationId: 'C1' }] },
    });
    await fetchConversationMessages(client(fetchFn), 'C1');
    const url = new URL(calls[0]);
    expect(url.searchParams.has('$orderby')).toBe(false);
    expect(url.searchParams.get('$filter')).toBe(`conversationId eq 'C1'`);
  });
});

describe('walkGraphDelta + accumulate', () => {
  it('follows @odata.nextLink across pages before returning the deltaLink', async () => {
    const pages: Record<string, GraphDeltaPage<Ms365DeltaMessage>> = {
      start: {
        value: [{ id: 'm1', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false }],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next',
      },
      'https://graph.microsoft.com/v1.0/next': {
        value: [{ id: 'm2', conversationId: 'C2', parentFolderId: 'inbox', isDraft: false }],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/final',
      },
    };
    const fetchFn = async (rawUrl: string) => {
      const key = rawUrl === 'https://graph.microsoft.com/v1.0/start' ? 'start' : rawUrl;
      const page = pages[key];
      if (!page) throw new Error(`no page for ${rawUrl}`);
      return {
        status: 200,
        statusText: '',
        headers: {},
        body: new TextEncoder().encode(JSON.stringify(page)),
      };
    };
    const c = client(fetchFn);
    const ids = new Set<string>();
    const deltaLink = await walkGraphDelta<Ms365DeltaMessage>(
      c,
      'https://graph.microsoft.com/v1.0/start',
      (page) => accumulate(page, ids),
    );
    expect(deltaLink).toBe('https://graph.microsoft.com/v1.0/final');
    expect([...ids].sort()).toEqual(['C1', 'C2']);
  });

  it('accumulate keeps every folder and drafts; only entries without a conversationId (e.g. @removed) are skipped', () => {
    const ids = new Set<string>();
    accumulate(
      {
        value: [
          { id: 'm1', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false },
          { id: 'm2', conversationId: 'C2', parentFolderId: 'inbox', isDraft: true },
          { id: 'm3', conversationId: 'C3', parentFolderId: 'JUNK', isDraft: false },
          { id: 'm4', conversationId: undefined, parentFolderId: 'inbox', isDraft: false },
        ],
      },
      ids,
    );
    expect([...ids]).toEqual(['C1', 'C2', 'C3']);
  });

  it('returns undefined when the signal is already aborted', async () => {
    const controller = new AbortController();
    controller.abort();
    const fetchFn = async () => {
      throw new Error('must not be called');
    };
    const result = await walkGraphDelta(
      client(fetchFn),
      'https://graph.microsoft.com/v1.0/start',
      () => {},
      controller.signal,
    );
    expect(result).toBeUndefined();
  });
});
