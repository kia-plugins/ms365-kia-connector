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
import type { LegacyMs365Cursor as Ms365Cursor } from '../cursor';
import type { Ms365ThreadItem } from '../to-document';
import { collect, graphFetch, graphMsg, instantClock, makeHost, makeSession } from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365ThreadItem>;

function makeSource(world: Parameters<typeof graphFetch>[0]) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  return { source, calls };
}

const liveCursor = (
  inboxDelta: string,
  sentDelta: string,
): Extract<Ms365Cursor, { phase: 'live' }> => ({
  phase: 'live',
  folders: { inbox: { delta: inboxDelta }, sentitems: { delta: sentDelta } },
});

describe('delta', () => {
  it('polls each folder, ingests affected conversations, advances deltaLinks', async () => {
    const { source } = makeSource({
      junkFolderId: 'JUNK',
      trashFolderId: 'TRASH',
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
    expect(batches[0].cursor).toEqual({
      phase: 'live',
      folders: {
        inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-next' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-next' },
      },
    });
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
      folders: { inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-next' } },
    });
  });

  it('recovers from a per-folder 410 syncStateNotFound by re-priming a 14-day window', async () => {
    const { source } = makeSource({
      custom: (url) => {
        if (
          url.pathname === '/v1.0/me/mailFolders/inbox/messages/delta' &&
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
          url.pathname === '/v1.0/me/mailFolders/inbox/messages/delta' &&
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
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=EXPIRED',
        'https://graph.microsoft.com/v1.0/sent-ok',
      )),
    )) as B[];

    expect(batches[0].cursor).toEqual({
      phase: 'live',
      folders: {
        inbox: { delta: 'https://graph.microsoft.com/v1.0/inbox-fresh' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-ok-next' },
      },
    });
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

  it('skips a folder with no delta cursor rather than throwing', async () => {
    const { source } = makeSource({
      urls: {
        'https://graph.microsoft.com/v1.0/sent-start': {
          value: [],
          '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/sent-next',
        },
      },
    });
    const { session } = makeSession();
    const cursor: Extract<Ms365Cursor, { phase: 'live' }> = {
      phase: 'live',
      folders: {
        inbox: { next: 'https://graph.microsoft.com/v1.0/inbox-mid-enumeration' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-start' },
      },
    };
    const batches = (await collect(source.pull(session, cursor))) as B[];
    expect(batches[0].cursor).toMatchObject({
      folders: {
        inbox: { next: 'https://graph.microsoft.com/v1.0/inbox-mid-enumeration' },
        sentitems: { delta: 'https://graph.microsoft.com/v1.0/sent-next' },
      },
    });
  });
});
