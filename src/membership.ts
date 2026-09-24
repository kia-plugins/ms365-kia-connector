/**
 * Which conversations have a message in a set of folders (spec §3.3/§3.4):
 * one paged `$select=conversationId` listing per folder. Shared by
 * `reconcile` (the tracked set) and `manageFolders` (leaving vs staying).
 * Pages are yielded as listed — duplicates across folders are fine, core's
 * reconcile staging ignores them and manageFolders collects into sets.
 */
import { GRAPH_BASE } from './graph-api';
import { GraphClient, statusOf } from './graph-client';

export async function* listConversationIds(
  client: GraphClient,
  folderIds: Iterable<string>,
  opts: { signal?: AbortSignal; warn?: (msg: string) => void; onRequest?: () => void } = {},
): AsyncGenerator<string[]> {
  for (const id of folderIds) {
    let url: string | undefined =
      `${GRAPH_BASE}/me/mailFolders/${encodeURIComponent(id)}/messages?$select=conversationId&$top=1000`;
    while (url) {
      if (opts.signal?.aborted) throw new Error('ms365: conversation listing aborted');
      let page: { value: Array<{ conversationId?: string | null }>; '@odata.nextLink'?: string };
      try {
        opts.onRequest?.();
        // eslint-disable-next-line no-await-in-loop
        page = await client.request(url);
      } catch (e) {
        // A folder deleted upstream since discovery (discovery never probes
        // leaves): its mail went with it, so listing it as empty is correct.
        if (statusOf(e) !== 404) throw e;
        opts.warn?.(`ms365: mail folder ${id} no longer exists — skipped`);
        break;
      }
      const ids = page.value.map((m) => m.conversationId).filter((c): c is string => !!c);
      if (ids.length) yield ids;
      url = page['@odata.nextLink'];
    }
  }
}
