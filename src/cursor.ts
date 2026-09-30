import { initialDeltaUrl, type FolderState } from './graph-api';

/** The two well-known folders a v1 cursor was keyed by. */
export type MailFolder = 'inbox' | 'sentitems';

/**
 * ms365's persisted Cursor (spec §3.2), keyed by Graph folder id over the
 * whole tracked folder tree.
 *
 * - `folders`: per tracked folder, either a `next` link still to walk (first
 *   enumeration, or a folder newly tracked) or the `delta` link live sweeps
 *   resume from.
 * - `pending`: conversationIds still to fetch and emit. `total` is the first
 *   backfill's fixed progress denominator.
 * - `retry`: conversations whose fetch failed (non-auth), with their
 *   consecutive-failure count. Never dropped: retried first on every pull
 *   until they succeed or resolve to a deletion.
 * - `phase` is the status label (`enumerate` → `ingest` → `live`); every
 *   pull runs the same steps regardless.
 */
export interface RetryEntry {
  id: string;
  n: number;
}

export interface Ms365Cursor {
  v: 2;
  phase: 'enumerate' | 'ingest' | 'live';
  folders: Record<string, FolderState>;
  pending: string[];
  total?: number;
  retry: RetryEntry[];
  /** The enumeration generation this cursor was built under. A cursor
   *  without the current one (loadCursor) restarts enumeration once —
   *  unchanged threads re-upsert with what the new generation adds.
   *  1 = attachment children (2.2.0, stored as `attachments: 1`);
   *  2 = reply targets (2.3.0). */
  rescan?: typeof RESCAN;
}

/** Current enumeration generation — bump to re-emit every conversation. */
export const RESCAN = 2;

/** v1: three phases over the two well-known folder names. */
export type LegacyMs365Cursor =
  | { phase: 'enumerate'; folders: Record<MailFolder, FolderState>; pending: string[] }
  | {
      phase: 'ingest';
      folders: Record<MailFolder, FolderState>;
      pending: string[];
      total: number;
    }
  | { phase: 'live'; folders: Record<MailFolder, FolderState> };

/** The ONE cursor surgery for a changed tracked set (every pull, and a
 *  scope Save): untracked folders' states are dropped, newly tracked
 *  folders start from their initial delta. `pending` and `retry` are kept —
 *  the emission gate decides their fate at fetch time. */
export function rescope(cursor: Ms365Cursor, tracked: ReadonlyMap<string, string>): Ms365Cursor {
  const folders: Record<string, FolderState> = {};
  for (const id of tracked.keys()) {
    folders[id] = cursor.folders[id] ?? { next: initialDeltaUrl(id) };
  }
  return { ...cursor, folders };
}
