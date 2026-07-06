/**
 * ms365 steady-state delta: one sweep of every mail folder's deltaLink per
 * `pull()` call — ported from legacy `ms365/delta.ts` `runDelta`, minus the
 * DB-backed deletion lookup (`deleteLocalConversation` there queried the
 * local documents table to also delete any attachment rows; since this port
 * never creates attachment sub-documents — see graph-api.ts — a deletion
 * here is always exactly one `email.thread` ExternalRef, and the ENGINE
 * archives it if (and only if) a matching document exists, so no `query`
 * cap / lookup is needed here at all).
 *
 * Ending after one sweep (rather than polling internally until aborted)
 * mirrors the gmail v2 port's `runDeltaSweep` choice: legacy's `runDelta`
 * was itself invoked once per Scheduler tick; the new engine's Cadence
 * (`every: '15m'`, see source.ts) plays that same external-timer role.
 */
import type { Batch, ExternalRef, Session } from './kiagent-contracts';
import type { Ms365Cursor } from './cursor';
import {
  fetchConversationMessages,
  MAIL_FOLDERS,
  primeDeltaUrl,
  resolveExcludedFolderIds,
  walkGraphDelta,
  type FolderState,
  type MailFolder,
  type Ms365DeltaMessage,
} from './graph-api';
import { accumulate } from './graph-api';
import { GraphClient, isAuthError, isSyncStateExpired } from './graph-client';
import { EMAIL_THREAD_DOCUMENT_TYPE, type Ms365ThreadItem } from './to-document';

const errText = (e: unknown): string => (e instanceof Error ? e.message : String(e));

/** How far back a re-primed delta window looks once a folder's deltaLink
 *  expires (410 syncStateNotFound) — verbatim from legacy `ms365/delta.ts`. */
const REPRIME_WINDOW_MS = 14 * 24 * 60 * 60 * 1000;

export async function* runDelta(
  client: GraphClient,
  session: Session,
  tenantKind: 'work' | 'personal',
  cursor: Extract<Ms365Cursor, { phase: 'live' }>,
): AsyncGenerator<Batch<Ms365Cursor, Ms365ThreadItem>> {
  const excluded = await resolveExcludedFolderIds(client);
  const newFolders: Record<MailFolder, FolderState> = { ...cursor.folders };
  const affected = new Set<string>();

  for (const folder of MAIL_FOLDERS) {
    if (session.signal.aborted) return;
    const prior = cursor.folders[folder];
    if (!prior || !('delta' in prior)) {
      session.log('warn', `ms365 delta: folder ${folder} has no delta cursor; skipping`);
      continue;
    }
    const walk = (startUrl: string) =>
      walkGraphDelta<Ms365DeltaMessage>(
        client,
        startUrl,
        (page) => accumulate(page, excluded, affected),
        session.signal,
      );
    let newDeltaLink: string | undefined;
    try {
      newDeltaLink = await walk(prior.delta);
    } catch (e) {
      if (isAuthError(e) || !isSyncStateExpired(e)) throw e;
      session.log(
        'warn',
        `ms365 delta: folder ${folder} deltaLink expired — re-priming ${REPRIME_WINDOW_MS / 86_400_000}-day window`,
      );
      const since = new Date(Date.now() - REPRIME_WINDOW_MS).toISOString();
      newDeltaLink = await walk(primeDeltaUrl(folder, since));
    }
    if (session.signal.aborted) return;
    if (newDeltaLink) newFolders[folder] = { delta: newDeltaLink };
  }

  const items: Ms365ThreadItem[] = [];
  const deletions: ExternalRef[] = [];
  for (const conversationId of affected) {
    if (session.signal.aborted) break;
    try {
      const messages = await fetchConversationMessages(client, conversationId);
      if (messages.length === 0) {
        deletions.push({ externalId: conversationId, type: EMAIL_THREAD_DOCUMENT_TYPE });
        continue;
      }
      items.push({ conversationId, messages, tenantKind });
    } catch (e) {
      if (isAuthError(e)) throw e;
      session.log(
        'warn',
        `ms365 delta: conversation ${conversationId} failed: ${errText(e)}`,
      );
    }
  }

  yield {
    phase: 'live',
    items,
    deletions: deletions.length ? deletions : undefined,
    cursor: { phase: 'live', folders: newFolders },
  };
}
