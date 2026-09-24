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
}

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

/** v1 → v2 without re-downloading anything: the `inbox`/`sentitems` keys
 *  become their resolved folder ids, every link, `pending` and `total`
 *  kept. */
export function migrateCursor(
  c: Ms365Cursor | LegacyMs365Cursor | null,
  wellKnownIds: Record<MailFolder, string>,
): Ms365Cursor | null {
  if (c === null || 'v' in c) return c;
  const folders: Record<string, FolderState> = {};
  for (const [name, state] of Object.entries(c.folders) as Array<[MailFolder, FolderState]>) {
    folders[wellKnownIds[name]] = state;
  }
  return {
    v: 2,
    phase: c.phase,
    folders,
    pending: 'pending' in c ? [...c.pending] : [],
    ...(c.phase === 'ingest' ? { total: c.total } : {}),
    retry: [],
  };
}

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
