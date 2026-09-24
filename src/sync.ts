/**
 * ms365 pull (spec §3.2): ONE generator for every phase over the tracked
 * folder tree. `phase` on the cursor is only a status label; each pull runs
 * the same steps:
 *
 *  1. walk every folder still in `next` state (first backfill, or a folder
 *     newly tracked) to its deltaLink, conversationIds → `pending`, with a
 *     checkpoint after every Graph page;
 *  2. drain `[...retry, ...pending]` through the fetch + emission gate;
 *  3. on a live pull, sweep every folder that was in `delta` state when the
 *     pull began (affected → `pending`), then drain again.
 *
 * Cursor + items commit together per yielded Batch, so a crash at any yield
 * resumes where the last commit left off; re-ingesting a conversation just
 * re-upserts the same document. A backfill pull (cursor null or not yet
 * live) ends by flipping the cursor to `live` without sweeping — the next
 * cadence tick sweeps, as before.
 */
import type { Batch, ExternalRef, Session } from '@kiagent/connector-sdk';
import type { Ms365Cursor, RetryEntry } from './cursor';
import {
  accumulate,
  fetchConversationMessages,
  primeDeltaUrl,
  walkGraphDelta,
  type FolderState,
  type GraphDeltaPage,
  type Ms365DeltaMessage,
} from './graph-api';
import { GraphClient, isAuthError, isSyncStateExpired } from './graph-client';
import type { GraphMessage } from './parser';
import type { ResolvedScope } from './scope';
import { EMAIL_THREAD_DOCUMENT_TYPE, type Ms365ThreadItem } from './to-document';

/** Two conversations fetched at once, v1 parity: an auth error rejecting
 *  either half of a pair propagates BEFORE the next pair is fetched. */
const INGEST_CONCURRENCY = 2;
/** A Batch (and cursor commit) every 25 processed conversations. */
const SAVE_PROGRESS_EVERY = 25;
/** From this many consecutive failures a retried id is logged at `warn`. */
const RETRY_WARN_AT = 5;
const RETRY_WARN_SIZE = 1000;
/** How far back a re-primed delta window looks once a folder's deltaLink
 *  expires (410 syncStateNotFound) — verbatim from legacy `ms365/delta.ts`. */
const REPRIME_WINDOW_MS = 14 * 24 * 60 * 60 * 1000;

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

type Gated = { item: Ms365ThreadItem } | { deletion: ExternalRef };

/** The emission gate (spec §3.2). A conversation is emitted only if one of
 *  its messages sits in the retention set — the tracked tree, or for a
 *  legacy account any folder at all — and is otherwise a deletion. A stale
 *  `pending`/`retry` entry therefore can never resurrect excluded mail. */
export function gate(
  conversationId: string,
  messages: GraphMessage[],
  scope: Pick<ResolvedScope, 'roots' | 'tracked' | 'legacy'>,
  tenantKind: 'work' | 'personal',
): Gated {
  const deletion = { deletion: { externalId: conversationId, type: EMAIL_THREAD_DOCUMENT_TYPE } };
  if (messages.length === 0) return deletion;
  if (scope.legacy) return { item: { conversationId, messages, tenantKind, scopeRootId: null } };
  const covering = new Set(
    messages.map((m) => scope.tracked.get(m.parentFolderId ?? '')).filter(Boolean),
  );
  const root = scope.roots.find((r) => covering.has(r.id));
  return root ? { item: { conversationId, messages, tenantKind, scopeRootId: root.id } } : deletion;
}

export async function* sync(
  client: GraphClient,
  session: Session,
  tenantKind: 'work' | 'personal',
  scope: ResolvedScope,
  start: Ms365Cursor,
): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>> {
  const backfill = start.phase !== 'live';
  const phase: 'backfill' | 'live' = backfill ? 'backfill' : 'live';
  const sweep = Object.keys(start.folders).filter((id) => 'delta' in start.folders[id]);
  let cur: Ms365Cursor = start;
  /** An id that failed this pull waits in `retry` for the next pull instead
   *  of being fetched again by the second drain. (One that SUCCEEDED is
   *  fetched again if the sweep saw it change — that is fresh content.) */
  const failedThisPull = new Set<string>();

  const batch = (items: Ms365ThreadItem[], deletions: ExternalRef[]) => ({
    phase,
    items,
    ...(deletions.length ? { deletions } : {}),
    cursor: cur,
    ...(backfill && cur.total !== undefined ? { estimateTotal: cur.total } : {}),
  });

  /** Fetch + gate every queued conversation. `force` yields once even when
   *  the queue is empty, so a cursor change made just before is committed. */
  async function* drain(force: boolean): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>> {
    const counts = new Map(cur.retry.map((r) => [r.id, r.n]));
    const queue = [...counts.keys(), ...cur.pending.filter((id) => !counts.has(id))].filter(
      (id) => !failedThisPull.has(id),
    );
    const held = cur.retry.filter((r) => failedThisPull.has(r.id));
    if (queue.length === 0) {
      if (force) yield batch([], []);
      return;
    }
    if (counts.size > RETRY_WARN_SIZE) {
      session.log('warn', `ms365: ${counts.size} conversations are waiting to be retried`);
    }
    const failed: RetryEntry[] = [];
    let items: Ms365ThreadItem[] = [];
    let deletions: ExternalRef[] = [];
    let since = 0;
    while (queue.length > 0) {
      if (session.signal.aborted) return;
      const pair = queue.splice(0, INGEST_CONCURRENCY);
      // eslint-disable-next-line no-await-in-loop
      const results = await Promise.all(
        pair.map(async (id) => {
          try {
            return { id, messages: await fetchConversationMessages(client, id) };
          } catch (e) {
            if (isAuthError(e)) throw e;
            return { id, error: e };
          }
        }),
      );
      for (const r of results) {
        if ('error' in r) {
          const n = (counts.get(r.id) ?? 0) + 1;
          failed.push({ id: r.id, n });
          failedThisPull.add(r.id);
          session.log(
            n >= RETRY_WARN_AT ? 'warn' : 'info',
            `ms365: conversation ${r.id} failed (${n} in a row), will retry: ${errText(r.error)}`,
          );
          continue;
        }
        const out = gate(r.id, r.messages ?? [], scope, tenantKind);
        if ('item' in out) items.push(out.item);
        else deletions.push(out.deletion);
      }
      since += pair.length;
      if (since >= SAVE_PROGRESS_EVERY || queue.length === 0) {
        cur = {
          ...cur,
          retry: [
            ...held,
            ...failed,
            ...queue.filter((id) => counts.has(id)).map((id) => ({ id, n: counts.get(id)! })),
          ],
          pending: queue.filter((id) => !counts.has(id)),
        };
        yield batch(items, deletions);
        items = [];
        deletions = [];
        since = 0;
      }
    }
  }

  // 1. Enumerate folders in `next` state, checkpointing every page.
  const pending = new Set(cur.pending);
  for (const [id, state] of Object.entries(cur.folders)) {
    if (!('next' in state)) continue;
    let url: string | undefined = state.next;
    while (url) {
      if (session.signal.aborted) return;
      // eslint-disable-next-line no-await-in-loop
      const page: GraphDeltaPage<Ms365DeltaMessage> = await client.request(url);
      accumulate(page, pending);
      const deltaLink = page['@odata.deltaLink'];
      const nextLink = page['@odata.nextLink'];
      if (!deltaLink && !nextLink) {
        throw new Error(`ms365: delta enumeration for folder ${id} ended without a nextLink or deltaLink`);
      }
      const next: FolderState = deltaLink ? { delta: deltaLink } : { next: nextLink! };
      cur = { ...cur, folders: { ...cur.folders, [id]: next }, pending: [...pending] };
      url = deltaLink ? undefined : nextLink;
      yield batch([], []);
    }
  }
  if (cur.phase === 'enumerate') {
    cur = { ...cur, phase: 'ingest', total: cur.pending.length };
    yield batch([], []);
  }

  // 2. Drain what is queued.
  yield* drain(false);
  if (session.signal.aborted) return;

  if (backfill) {
    const { total: _done, ...rest } = cur;
    cur = { ...rest, phase: 'live' };
    yield { phase: 'live', items: [], cursor: cur };
    return;
  }

  // 3. Live sweep of the folders that were already live.
  const affected = new Set<string>();
  const folders = { ...cur.folders };
  for (const id of sweep) {
    if (session.signal.aborted) return;
    const prior = folders[id];
    if (!prior || !('delta' in prior)) continue;
    const walk = (startUrl: string) =>
      walkGraphDelta<Ms365DeltaMessage>(
        client,
        startUrl,
        (page) => accumulate(page, affected),
        session.signal,
      );
    let link: string | undefined;
    try {
      // eslint-disable-next-line no-await-in-loop
      link = await walk(prior.delta);
    } catch (e) {
      if (isAuthError(e) || !isSyncStateExpired(e)) throw e;
      session.log(
        'warn',
        `ms365 delta: folder ${id} deltaLink expired — re-priming ${REPRIME_WINDOW_MS / 86_400_000}-day window`,
      );
      const since = new Date(Date.now() - REPRIME_WINDOW_MS).toISOString();
      // eslint-disable-next-line no-await-in-loop
      link = await walk(primeDeltaUrl(id, since));
    }
    if (session.signal.aborted) return;
    if (link) folders[id] = { delta: link };
  }
  // The advanced links commit with the affected ids already in `pending`,
  // so nothing the sweep saw can be lost to a crash mid-drain.
  cur = {
    ...cur,
    folders,
    pending: [...new Set([...cur.pending, ...affected])],
  };
  yield* drain(true);
}
