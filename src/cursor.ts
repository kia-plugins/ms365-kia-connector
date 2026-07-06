import type { FolderState, MailFolder } from './graph-api';

/**
 * ms365's persisted Cursor shape — a three-phase state machine ported from
 * legacy's per-folder `Ms365Cursor` (`ms365/backfill.ts`), reshaped onto a
 * single `Source<Cursor, Item>.pull()` async generator (google-docs /
 * gmail v2 idiom) instead of legacy's separate `runBackfill`/`runDelta`
 * entry points invoked by an external scheduler.
 *
 * - `enumerate`: walking each MAIL_FOLDER's `/messages/delta` feed to
 *   collect the set of non-draft, non-excluded-folder conversationIds
 *   (`pending`), until every folder has captured a final deltaLink.
 * - `ingest`: fetching each pending conversationId's full message set and
 *   emitting a thread document per non-empty conversation. `total` is the
 *   enumeration's conversation count — a fixed backfill-progress
 *   denominator, captured once when `enumerate` completes.
 * - `live`: steady-state delta polling — one sweep of every folder's
 *   deltaLink per `pull()` call (cadence-driven, like the gmail v2 port's
 *   single-sweep-per-tick design; see source.ts).
 */
export type Ms365Cursor =
  | { phase: 'enumerate'; folders: Record<MailFolder, FolderState>; pending: string[] }
  | {
      phase: 'ingest';
      folders: Record<MailFolder, FolderState>;
      pending: string[];
      total: number;
    }
  | { phase: 'live'; folders: Record<MailFolder, FolderState> };
