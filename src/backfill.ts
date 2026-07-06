/**
 * ms365 backfill: per-folder delta enumeration (v1 `enumerateConversations`)
 * followed by per-conversation ingest (v1 `ingestConversations`), reshaped
 * as ONE async generator yielding `Batch`es — cursor + items commit
 * together in the same engine transaction, so a crash at any yield boundary
 * resumes exactly where the last committed batch left off (idempotent:
 * re-ingesting an already-emitted conversationId just re-upserts the same
 * document).
 */
import type { Batch, Session } from './kiagent-contracts';
import type { Ms365Cursor } from './cursor';
import {
  accumulate,
  fetchConversationMessages,
  initialDeltaUrl,
  MAIL_FOLDERS,
  resolveExcludedFolderIds,
  walkGraphDelta,
  type FolderState,
  type MailFolder,
  type Ms365DeltaMessage,
} from './graph-api';
import { GraphClient, isAuthError } from './graph-client';
import type { Ms365ThreadItem } from './to-document';

/** v1 `INGEST_CONCURRENCY` — two conversations fetched at once, matching
 *  legacy's exact concurrency/failure semantics: an auth error rejecting
 *  either half of a pair propagates BEFORE the next pair is ever fetched
 *  (see the "rethrows an auth error" legacy test, ported in
 *  `__tests__/backfill.test.ts`). */
const INGEST_CONCURRENCY = 2;
/** v1 `SAVE_PROGRESS_EVERY` — batches (and thus cursor commits) are yielded
 *  every 25 processed conversations rather than after every pair, so a
 *  restart mid-run re-does at most 25 conversations' worth of (idempotent,
 *  cheap) work. */
const SAVE_PROGRESS_EVERY = 25;

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

interface EnumerateResult {
  folders: Record<MailFolder, FolderState>;
  pending: string[];
}

/**
 * Runs (or resumes) the per-folder delta enumeration. Yields an interim
 * `backfill` batch (no items) after each folder finishes enumerating, so a
 * crash mid-enumeration resumes from the last folder boundary rather than
 * restarting every folder. Returns the fully-enumerated folder states and
 * pending conversationId list once every folder has captured its deltaLink.
 */
async function* enumerate(
  client: GraphClient,
  session: Session,
  resume: Extract<Ms365Cursor, { phase: 'enumerate' }> | undefined,
): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>, EnumerateResult> {
  const excluded = await resolveExcludedFolderIds(client);
  const conversationIds = new Set<string>(resume?.pending ?? []);
  const folders = {} as Record<MailFolder, FolderState>;
  for (const f of MAIL_FOLDERS) {
    folders[f] = resume?.folders[f] ?? { next: initialDeltaUrl(f) };
  }

  for (const folder of MAIL_FOLDERS) {
    if (session.signal.aborted) return { folders, pending: [...conversationIds] };
    const state = folders[folder];
    if ('delta' in state) continue; // resumed folder already fully enumerated
    const deltaLink = await walkGraphDelta<Ms365DeltaMessage>(
      client,
      state.next,
      (page) => {
        accumulate(page, excluded, conversationIds);
        if (!page['@odata.deltaLink'] && page['@odata.nextLink']) {
          folders[folder] = { next: page['@odata.nextLink'] };
        }
      },
      session.signal,
    );
    if (session.signal.aborted) return { folders, pending: [...conversationIds] };
    if (!deltaLink) {
      throw new Error(
        `ms365: delta enumeration for folder ${folder} ended without a nextLink or deltaLink`,
      );
    }
    folders[folder] = { delta: deltaLink };
    yield {
      phase: 'backfill',
      items: [],
      cursor: { phase: 'enumerate', folders, pending: [...conversationIds] },
    };
  }
  return { folders, pending: [...conversationIds] };
}

/**
 * Fetches + emits each pending conversation, `INGEST_CONCURRENCY` at a
 * time, yielding a `Batch` every `SAVE_PROGRESS_EVERY` processed (or when
 * `pending` drains). A conversation that resolves to zero messages is
 * skipped silently (legacy's "empty conversation" case — nothing to
 * archive during backfill, since nothing was ever committed for it). Any
 * OTHER per-conversation failure is logged and skipped; an auth error
 * propagates immediately, before the next pair is ever fetched.
 */
async function* ingest(
  client: GraphClient,
  session: Session,
  tenantKind: 'work' | 'personal',
  folders: Record<MailFolder, FolderState>,
  pending: string[],
  total: number,
): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>> {
  let buffer: Ms365ThreadItem[] = [];
  let sinceYield = 0;
  while (pending.length > 0) {
    if (session.signal.aborted) return;
    const batch = pending.splice(0, INGEST_CONCURRENCY);
    const results = await Promise.all(
      batch.map(async (conversationId): Promise<Ms365ThreadItem | null> => {
        try {
          const messages = await fetchConversationMessages(client, conversationId);
          return messages.length > 0 ? { conversationId, messages, tenantKind } : null;
        } catch (e) {
          if (isAuthError(e)) throw e;
          session.log(
            'warn',
            `ms365 backfill: conversation ${conversationId} failed: ${errText(e)}`,
          );
          return null;
        }
      }),
    );
    for (const r of results) if (r) buffer.push(r);
    sinceYield += batch.length;

    if (sinceYield >= SAVE_PROGRESS_EVERY || pending.length === 0) {
      yield {
        phase: 'backfill',
        items: buffer,
        cursor: { phase: 'ingest', folders, pending: [...pending], total },
        estimateTotal: total,
      };
      buffer = [];
      sinceYield = 0;
    }
  }
}

export async function* runBackfill(
  client: GraphClient,
  session: Session,
  tenantKind: 'work' | 'personal',
  cursor: Ms365Cursor | null,
): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>> {
  let folders: Record<MailFolder, FolderState>;
  let pending: string[];
  let total: number;

  if (cursor?.phase === 'ingest') {
    folders = cursor.folders;
    pending = [...cursor.pending];
    total = cursor.total;
  } else {
    const resume = cursor?.phase === 'enumerate' ? cursor : undefined;
    const result = yield* enumerate(client, session, resume);
    if (session.signal.aborted) return;
    folders = result.folders;
    pending = result.pending;
    total = pending.length;
    yield {
      phase: 'backfill',
      items: [],
      cursor: { phase: 'ingest', folders, pending: [...pending], total },
    };
  }

  yield* ingest(client, session, tenantKind, folders, pending, total);
  if (session.signal.aborted) return;
  yield { phase: 'live', items: [], cursor: { phase: 'live', folders } };
}
