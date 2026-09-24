/**
 * Microsoft Graph mail endpoints used by this connector: the folder-delta
 * enumeration walk (`walkGraphDelta`/`accumulate`) and the full-conversation
 * message fetch — ported from the legacy
 * v1 repo's `ms365/client.ts` + `ms-shared/walk-delta.ts`, reshaped onto
 * `GraphClient` (host `net.fetch`) instead of the legacy positional
 * `graphFetch(url, getToken)` function.
 */
import { GraphClient } from './graph-client';
import type { GraphMessage } from './parser';

/** Public, well-known Microsoft Graph v1.0 base URL — not a credential. */
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

/**
 * Personal Microsoft accounts (MSA tenant) cannot delta-query /me/messages —
 * the endpoint only supports change tracking when scoped to a specific mail
 * folder, and it is NOT recursive: every tracked folder (subfolders
 * included) is delta-queried on its own (see folders.ts).
 */

/** Per-folder paging state: either we still need to fetch `next` (a
 *  nextLink URL produced by Graph), or we have a final `delta` URL that
 *  becomes the starting point for live delta polling. */
export type FolderState = { next: string } | { delta: string };

export const SELECT_FIELDS = 'id,conversationId,parentFolderId,isDraft';

/** `folder` is a Graph folder id (or a well-known name). */
export function initialDeltaUrl(folder: string): string {
  return `${GRAPH_BASE}/me/mailFolders/${encodeURIComponent(folder)}/messages/delta?$select=${SELECT_FIELDS}&$top=100`;
}

export function primeDeltaUrl(folder: string, sinceIso: string): string {
  const filter = `receivedDateTime ge ${sinceIso}`;
  return (
    `${GRAPH_BASE}/me/mailFolders/${encodeURIComponent(folder)}/messages/delta` +
    `?$select=${SELECT_FIELDS}` +
    `&$filter=${encodeURIComponent(filter)}` +
    `&$top=100`
  );
}

/** Shape of one page of Graph's OData delta feed. */
export interface GraphDeltaPage<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

/**
 * Walks an OData delta feed: GET url → await onPage(page) → follow
 * `@odata.nextLink` until a page carries `@odata.deltaLink`.
 *
 * Returns that final deltaLink URL, or `undefined` when the walk stopped
 * without one because `signal` aborted (checked before every fetch).
 * Callers that persist a resume cursor mid-walk read `@odata.nextLink`
 * themselves inside `onPage`, before the next fetch happens.
 */
export async function walkGraphDelta<T>(
  client: GraphClient,
  startUrl: string,
  onPage: (page: GraphDeltaPage<T>) => void | Promise<void>,
  signal?: AbortSignal,
): Promise<string | undefined> {
  let url: string | undefined = startUrl;
  while (url) {
    if (signal?.aborted) return undefined;
    const page: GraphDeltaPage<T> = await client.request<GraphDeltaPage<T>>(url);
    await onPage(page);
    if (page['@odata.deltaLink']) return page['@odata.deltaLink'];
    url = page['@odata.nextLink'];
  }
  return undefined;
}

export interface Ms365DeltaMessage {
  id?: string;
  conversationId?: string | null;
  parentFolderId?: string;
  isDraft?: boolean;
  '@removed'?: { reason?: string };
}

/** Folds one delta page's conversationIds into `into`. Eligibility is
 *  folder membership only (spec §3.2): drafts count, and every page comes
 *  from a tracked folder. `@removed` entries carry no conversationId and
 *  fall out here — moves out of scope are reconcile's job. */
export function accumulate(page: GraphDeltaPage<Ms365DeltaMessage>, into: Set<string>): void {
  for (const m of page.value) {
    if (m.conversationId) into.add(m.conversationId);
  }
}

// $select for the full conversation fetch. Deliberately excludes
// `attachments` — this connector does not ingest attachment bytes.
const CONV_SELECT =
  'id,subject,from,toRecipients,ccRecipients,bccRecipients,' +
  'receivedDateTime,internetMessageHeaders,body,bodyPreview,' +
  'hasAttachments,parentFolderId,internetMessageId,conversationId';

/**
 * Fetches every message in one conversation, oldest first. No `$orderby`:
 * combining it with a conversationId filter trips Graph's
 * "InefficientFilter" 400 on personal MS accounts — sort client-side
 * instead (a single thread is small enough that the cost is irrelevant).
 * Verbatim from legacy `ms365/client.ts`.
 */
export async function fetchConversationMessages(
  client: GraphClient,
  conversationId: string,
): Promise<GraphMessage[]> {
  const messages: GraphMessage[] = [];
  const initial = new URL(`${GRAPH_BASE}/me/messages`);
  initial.searchParams.set('$filter', `conversationId eq '${conversationId}'`);
  initial.searchParams.set('$select', CONV_SELECT);
  initial.searchParams.set('$top', '50');
  let url: string | undefined = initial.toString();
  while (url) {
    const page: { value: GraphMessage[]; '@odata.nextLink'?: string } =
      await client.request(url, {
        extraHeaders: { prefer: 'outlook.body-content-type="text"' },
      });
    messages.push(...page.value);
    url = page['@odata.nextLink'];
  }
  messages.sort((a, b) => {
    const ta = Date.parse(a.receivedDateTime ?? '') || 0;
    const tb = Date.parse(b.receivedDateTime ?? '') || 0;
    return ta - tb;
  });
  return messages;
}
