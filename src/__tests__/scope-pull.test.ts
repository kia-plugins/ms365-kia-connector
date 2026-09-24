/**
 * Pull over the tracked folder tree (spec §3.2): discovery every pull, new
 * folders enumerated inside `live`, the membership emission gate, durable
 * retry, and the lossless legacy-cursor path.
 */
import type { Batch } from '@kiagent/connector-sdk';
import { createMs365Source } from '../source';
import { initialDeltaUrl } from '../graph-api';
import type { Ms365Cursor } from '../cursor';
import { EMAIL_THREAD_DOCUMENT_TYPE, type Ms365ThreadItem } from '../to-document';
import {
  collect,
  DEFAULT_FOLDERS,
  graphFetch,
  graphMsg,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
  type GraphWorld,
} from '../testing/harness';

type B = Batch<Ms365Cursor, Ms365ThreadItem>;

const G = 'https://graph.microsoft.com/v1.0';
const INBOX_ROOT = [{ id: 'INBOX-ID', name: 'Inbox' }];

function run(world: GraphWorld, config: Record<string, unknown>, cursor: unknown) {
  const { fetchFn, calls } = graphFetch(world);
  const source = createMs365Source(makeHost(fetchFn), instantClock);
  const { session, logs } = makeSession({ config: { tenantKind: 'personal', ...config } });
  return {
    calls,
    logs,
    batches: collect(source.pull(session, cursor as Ms365Cursor)) as Promise<B[]>,
  };
}

const live = (folders: Ms365Cursor['folders'], extra: Partial<Ms365Cursor> = {}): Ms365Cursor => ({
  v: 2,
  phase: 'live',
  folders,
  pending: [],
  retry: [],
  ...extra,
});

const items = (bs: B[]) => bs.flatMap((b) => b.items);
const deletions = (bs: B[]) => bs.flatMap((b) => b.deletions ?? []);
const last = (bs: B[]) => bs[bs.length - 1];

describe('pull over the tracked folder tree', () => {
  it('a new subfolder under Inbox is enumerated and ingested, stamped with the Inbox root', async () => {
    const { batches } = run(
      {
        folders: {
          ...DEFAULT_FOLDERS,
          children: { 'INBOX-ID': [{ id: 'SUB', displayName: 'Sub', childFolderCount: 0 }] },
        },
        urls: {
          [`${G}/inbox-d`]: { value: [], '@odata.deltaLink': `${G}/inbox-d2` },
          [initialDeltaUrl('SUB')]: {
            value: [{ id: 's1', conversationId: 'CS', parentFolderId: 'SUB' }],
            '@odata.deltaLink': `${G}/sub-d`,
          },
        },
        conversations: { CS: [graphMsg({ conversationId: 'CS', parentFolderId: 'SUB' })] },
      },
      { folderRoots: INBOX_ROOT },
      live({ 'INBOX-ID': { delta: `${G}/inbox-d` } }),
    );
    const bs = await batches;
    expect(bs.every((b) => b.phase === 'live')).toBe(true);
    expect(items(bs).map((i) => [i.conversationId, i.scopeRootId])).toEqual([['CS', 'INBOX-ID']]);
    expect(last(bs).cursor.folders).toEqual({
      'INBOX-ID': { delta: `${G}/inbox-d2` },
      SUB: { delta: `${G}/sub-d` },
    });
  });

  it('a pending conversation whose messages all left the tracked set is a deletion, not an item', async () => {
    const { batches } = run(
      {
        urls: { [`${G}/inbox-d`]: { value: [], '@odata.deltaLink': `${G}/inbox-d2` } },
        conversations: { CX: [graphMsg({ conversationId: 'CX', parentFolderId: 'DELETED-ID' })] },
      },
      { folderRoots: INBOX_ROOT },
      live({ 'INBOX-ID': { delta: `${G}/inbox-d` } }, { pending: ['CX'] }),
    );
    const bs = await batches;
    expect(items(bs)).toEqual([]);
    expect(deletions(bs)).toEqual([{ externalId: 'CX', type: EMAIL_THREAD_DOCUMENT_TYPE }]);
    expect(last(bs).cursor.pending).toEqual([]);
  });

  it('legacy account: a conversation living only in Deleted Items is KEPT (no stamp)', async () => {
    const { batches } = run(
      {
        urls: {
          [`${G}/inbox-d`]: { value: [], '@odata.deltaLink': `${G}/inbox-d2` },
          [`${G}/sent-d`]: { value: [], '@odata.deltaLink': `${G}/sent-d2` },
        },
        conversations: { CX: [graphMsg({ conversationId: 'CX', parentFolderId: 'DELETED-ID' })] },
      },
      {},
      live(
        { 'INBOX-ID': { delta: `${G}/inbox-d` }, 'SENT-ID': { delta: `${G}/sent-d` } },
        { pending: ['CX'] },
      ),
    );
    const bs = await batches;
    expect(items(bs).map((i) => [i.conversationId, i.scopeRootId])).toEqual([['CX', null]]);
    expect(deletions(bs)).toEqual([]);
  });

  it('a draft in a tracked folder is enumerated', async () => {
    const { batches } = run(
      {
        urls: {
          [initialDeltaUrl('INBOX-ID')]: {
            value: [{ id: 'd1', conversationId: 'CD', parentFolderId: 'INBOX-ID', isDraft: true }],
            '@odata.deltaLink': `${G}/inbox-d`,
          },
        },
        conversations: { CD: [graphMsg({ conversationId: 'CD', parentFolderId: 'INBOX-ID' })] },
      },
      { folderRoots: INBOX_ROOT },
      null,
    );
    const bs = await batches;
    expect(items(bs).map((i) => i.conversationId)).toEqual(['CD']);
    expect(last(bs).cursor.phase).toBe('live');
  });

  it('a legacy v1 live cursor keeps its delta links: no initial delta for Inbox or Sent', async () => {
    const { batches, calls } = run(
      {
        urls: {
          [`${G}/legacy-inbox`]: { value: [], '@odata.deltaLink': `${G}/inbox-d2` },
          [`${G}/legacy-sent`]: { value: [], '@odata.deltaLink': `${G}/sent-d2` },
        },
      },
      {},
      { phase: 'live', folders: { inbox: { delta: `${G}/legacy-inbox` }, sentitems: { delta: `${G}/legacy-sent` } } },
    );
    const bs = await batches;
    expect(calls).toContain(`${G}/legacy-inbox`);
    expect(calls).toContain(`${G}/legacy-sent`);
    expect(calls.some((u) => u.includes('/messages/delta?$select'))).toBe(false);
    expect(last(bs).cursor).toEqual(
      live({ 'INBOX-ID': { delta: `${G}/inbox-d2` }, 'SENT-ID': { delta: `${G}/sent-d2` } }),
    );
  });
});

describe('durable retry', () => {
  const failing = (id: string): GraphWorld['custom'] => (url) =>
    url.pathname === '/v1.0/me/messages' &&
    (url.searchParams.get('$filter') ?? '').includes(`'${id}'`)
      ? jsonRes(500, {})
      : undefined;
  const world = (custom?: GraphWorld['custom']): GraphWorld => ({
    urls: { [`${G}/inbox-d`]: { value: [], '@odata.deltaLink': `${G}/inbox-d` } },
    conversations: { BAD: [graphMsg({ conversationId: 'BAD', parentFolderId: 'INBOX-ID' })] },
    custom,
  });
  const start = (extra: Partial<Ms365Cursor>) => live({ 'INBOX-ID': { delta: `${G}/inbox-d` } }, extra);

  it('a failed fetch is committed to retry with n:1; the next pull succeeds and clears it', async () => {
    const first = await run(world(failing('BAD')), { folderRoots: INBOX_ROOT }, start({ pending: ['BAD'] }))
      .batches;
    expect(items(first)).toEqual([]);
    expect(last(first).cursor.retry).toEqual([{ id: 'BAD', n: 1 }]);
    expect(last(first).cursor.pending).toEqual([]);

    const second = await run(world(), { folderRoots: INBOX_ROOT }, last(first).cursor).batches;
    expect(items(second).map((i) => i.conversationId)).toEqual(['BAD']);
    expect(last(second).cursor.retry).toEqual([]);
  }, 30_000);

  it('from 5 consecutive failures the id is logged at warn, and it stays in retry', async () => {
    const r = run(world(failing('BAD')), { folderRoots: INBOX_ROOT }, start({ retry: [{ id: 'BAD', n: 5 }] }));
    const bs = await r.batches;
    expect(last(bs).cursor.retry).toEqual([{ id: 'BAD', n: 6 }]);
    expect(r.logs.some((l) => l.level === 'warn' && l.msg.includes('BAD'))).toBe(true);
  }, 30_000);
});
