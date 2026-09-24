/**
 * Steady-state delta suite. Ported from legacy `src/__tests__/ms365-delta.test.ts`
 * (runDelta): polls each folder, ingests affected conversations, advances
 * deltaLinks, recovers from a 410 syncStateNotFound by re-priming a 14-day
 * window, and (new in v2 — legacy deleted straight from its own DB) surfaces
 * a zero-message conversation as a `Batch.deletions` ExternalRef instead.
 */
import { createMs365Source } from '../source';
import { EMAIL_THREAD_DOCUMENT_TYPE } from '../to-document';
import type { Batch } from '@kiagent/connector-sdk';
import type { Ms365Cursor } from '../cursor';
import type { Ms365ThreadItem } from '../to-document';
import { collect, graphFetch, graphMsg, instantClock, makeHost, makeSession } from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365ThreadItem>;

function makeSource(world: Parameters<typeof graphFetch>[0]) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  return { source, calls };
}

const liveCursor = (inboxDelta: string, sentDelta: string): Ms365Cursor => ({
  v: 2,
  phase: 'live',
  folders: { 'INBOX-ID': { delta: inboxDelta }, 'SENT-ID': { delta: sentDelta } },
  pending: [],
  retry: [],
});

describe('delta', () => {
  it('polls each folder, ingests affected conversations, advances deltaLinks', async () => {
    const { source } = makeSource({
      urls: {
        'https://graph.microsoft.com/v1.0/inbox-start': {
          value: [{ id: 'm-changed', conversationId: 'CA', parentFolderId: 'inbox', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-next',
        },
        'https://graph.microsoft.com/v1.0/sent-start': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-next',
        },
      },
      conversations: { CA: [graphMsg({ id: 'm-changed', conversationId: 'CA', subject: 'updated' })] },
    });
    const { session } = makeSession();

    const batches = (await collect(
      source.pull(session, liveCursor(
        'https://graph.microsoft.com/v1.0/inbox-start',
        'https://graph.microsoft.com/v1.0/sent-start',
      )),
    )) as B[];

    expect(batches).toHaveLength(1);
    expect(batches[0].phase).toBe('live');
    expect(batches[0].items.map((i) => i.conversationId)).toEqual(['CA']);
    expect(batches[0].cursor).toEqual(
      liveCursor('https://graph.microsoft.com/v1.0/inbox-next', 'https://graph.microsoft.com/v1.0/sent-next'),
    );
  });

  it('follows @odata.nextLink across pages before capturing the deltaLink', async () => {
    const { source } = makeSource({
      urls: {
        'https://graph.microsoft.com/v1.0/inbox-start': {
          value: [{ id: 'm1', conversationId: 'C1', parentFolderId: 'inbox', isDraft: false }],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/inbox-page2',
        },
        'https://graph.microsoft.com/v1.0/inbox-page2': {
          value: [{ id: 'm2', conversationId: 'C2', parentFolderId: 'inbox', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-next',
        },
        'https://graph.microsoft.com/v1.0/sent-start': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-next',
        },
      },
      conversations: {
        C1: [graphMsg({ conversationId: 'C1', subject: 'subj C1' })],
        C2: [graphMsg({ conversationId: 'C2', subject: 'subj C2' })],
      },
    });
    const { session } = makeSession();
    const batches = (await collect(
      source.pull(session, liveCursor(
        'https://graph.microsoft.com/v1.0/inbox-start',
        'https://graph.microsoft.com/v1.0/sent-start',
      )),
    )) as B[];
    expect(batches[0].items.map((i) => i.conversationId).sort()).toEqual(['C1', 'C2']);
    expect(batches[0].cursor).toMatchObject({
      folders: { 'INBOX-ID': { delta: 'https://graph.microsoft.com/v1.0/inbox-next' } },
    });
  });

  it('recovers from a per-folder 410 syncStateNotFound by re-priming a 14-day window', async () => {
    const { source } = makeSource({
      custom: (url) => {
        if (
          url.pathname === '/v1.0/me/mailFolders/INBOX-ID/messages/delta' &&
          url.searchParams.get('$deltatoken') === 'EXPIRED'
        ) {
          return {
            status: 410,
            statusText: '',
            headers: {},
            body: new TextEncoder().encode(
              JSON.stringify({ error: { code: 'syncStateNotFound', message: 'gone' } }),
            ),
          };
        }
        if (
          url.pathname === '/v1.0/me/mailFolders/INBOX-ID/messages/delta' &&
          (url.searchParams.get('$filter') ?? '').startsWith('receivedDateTime ge')
        ) {
          return {
            status: 200,
            statusText: '',
            headers: {},
            body: new TextEncoder().encode(
              JSON.stringify({
                value: [{ id: 'mR', conversationId: 'CR', parentFolderId: 'inbox', isDraft: false }],
                '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-fresh',
              }),
            ),
          };
        }
        return undefined;
      },
      urls: {
        'https://graph.microsoft.com/v1.0/sent-ok': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-ok-next',
        },
      },
      conversations: { CR: [graphMsg({ conversationId: 'CR', subject: 'recent' })] },
    });
    const { session } = makeSession();

    const batches = (await collect(
      source.pull(session, liveCursor(
        'https://graph.microsoft.com/v1.0/me/mailFolders/INBOX-ID/messages/delta?$deltatoken=EXPIRED',
        'https://graph.microsoft.com/v1.0/sent-ok',
      )),
    )) as B[];

    expect(batches[0].cursor).toEqual(
      liveCursor('https://graph.microsoft.com/v1.0/inbox-fresh', 'https://graph.microsoft.com/v1.0/sent-ok-next'),
    );
    expect(batches[0].items.map((i) => i.conversationId)).toEqual(['CR']);
  });

  it('surfaces a zero-message conversation as a deletion, not an item', async () => {
    const { source } = makeSource({
      urls: {
        'https://graph.microsoft.com/v1.0/inbox-start': {
          value: [{ id: 'mD', conversationId: 'DELETED', parentFolderId: 'inbox', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-next',
        },
        'https://graph.microsoft.com/v1.0/sent-start': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-next',
        },
      },
      conversations: { DELETED: [] },
    });
    const { session } = makeSession();
    const batches = (await collect(
      source.pull(session, liveCursor(
        'https://graph.microsoft.com/v1.0/inbox-start',
        'https://graph.microsoft.com/v1.0/sent-start',
      )),
    )) as B[];
    expect(batches[0].items).toEqual([]);
    expect(batches[0].deletions).toEqual([
      { externalId: 'DELETED', type: EMAIL_THREAD_DOCUMENT_TYPE },
    ]);
  });

  it('a live folder still in `next` state is enumerated inside the live pull (it used to be skipped)', async () => {
    const { source } = makeSource({
      urls: {
        'https://graph.microsoft.com/v1.0/inbox-mid-enumeration': {
          value: [{ id: 'mN', conversationId: 'CN', parentFolderId: 'INBOX-ID', isDraft: false }],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/inbox-next',
        },
        'https://graph.microsoft.com/v1.0/sent-start': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-next',
        },
      },
      conversations: { CN: [graphMsg({ conversationId: 'CN' })] },
    });
    const { session } = makeSession();
    const cursor: Ms365Cursor = {
      ...liveCursor('unused', 'https://graph.microsoft.com/v1.0/sent-start'),
      folders: {
        'INBOX-ID': { next: 'https://graph.microsoft.com/v1.0/inbox-mid-enumeration' },
        'SENT-ID': { delta: 'https://graph.microsoft.com/v1.0/sent-start' },
      },
    };
    const batches = (await collect(source.pull(session, cursor))) as B[];
    expect(batches.every((b) => b.phase === 'live')).toBe(true);
    expect(batches.flatMap((b) => b.items.map((i) => i.conversationId))).toEqual(['CN']);
    expect(batches[batches.length - 1].cursor).toEqual(
      liveCursor('https://graph.microsoft.com/v1.0/inbox-next', 'https://graph.microsoft.com/v1.0/sent-next'),
    );
  });
});
