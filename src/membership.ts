/**
 * Which conversations have a message in a set of folders (spec §3.3/§3.4):
 * one paged `$select=conversationId` listing per folder. Shared by
 * `reconcile` (the tracked set) and `manageFolders` (leaving vs staying).
 * Pages are yielded as listed — duplicates across folders are fine, core's
 * reconcile staging ignores them and manageFolders collects into sets.
 */
import { GRAPH_BASE } from './graph-api';
import { GraphClient } from './graph-client';

export async function* listConversationIds(
  client: GraphClient,
  folderIds: Iterable<string>,
  opts: { signal?: AbortSignal; onRequest?: () => void } = {},
): AsyncGenerator<string[]> {
  for (const id of folderIds) {
    let url: string | undefined =
      `${GRAPH_BASE}/me/mailFolders/${encodeURIComponent(id)}/messages?$select=conversationId&$top=1000`;
    while (url) {
      if (opts.signal?.aborted) throw new Error('ms365: conversation listing aborted');
      // Every folder here was listed by discovery moments ago, so ANY error
      // — a 404 included — fails the listing: read as empty, a folder would
      // archive its whole index with no re-enumeration to bring it back.
      // A genuinely deleted folder is dropped at discovery next cycle.
      opts.onRequest?.();
      // eslint-disable-next-line no-await-in-loop
      const page: { value: Array<{ conversationId?: string | null }>; '@odata.nextLink'?: string } =
        await client.request(url);
      const ids = page.value.map((m) => m.conversationId).filter((c): c is string => !!c);
      if (ids.length) yield ids;
      url = page['@odata.nextLink'];
    }
  }
}
