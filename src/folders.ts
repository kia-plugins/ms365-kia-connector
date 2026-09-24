/**
 * Outlook mail folder discovery. Graph's message delta is per folder and NOT
 * recursive, so a selected folder's subfolders must be enumerated here and
 * delta-queried one by one. Discovery fails closed: any request error
 * propagates, because a partial folder map would read as "those folders'
 * mail is gone" to everything downstream.
 */
import { GRAPH_BASE } from './graph-api';
import { GraphClient, statusOf } from './graph-client';

export interface MailFolderNode {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount: number;
  wellKnown?: string;
}

const FOLDER_SELECT = 'id,displayName,parentFolderId,childFolderCount';

async function listPaged(client: GraphClient, first: string): Promise<MailFolderNode[]> {
  const out: MailFolderNode[] = [];
  let url: string | undefined = first;
  while (url) {
    const page: { value: MailFolderNode[]; '@odata.nextLink'?: string } =
      await client.request(url);
    out.push(...page.value);
    url = page['@odata.nextLink'];
  }
  return out;
}

/** Well-known folders by name (`inbox`, `sentitems`, `archive`, …). A name
 *  the mailbox does not have (404 — e.g. no Archive folder) is omitted;
 *  any other error propagates. */
export async function resolveWellKnown(
  client: GraphClient,
  names: string[],
): Promise<Record<string, MailFolderNode>> {
  const out: Record<string, MailFolderNode> = {};
  for (const name of names) {
    try {
      // eslint-disable-next-line no-await-in-loop
      const node = await client.request<MailFolderNode>(
        `${GRAPH_BASE}/me/mailFolders/${name}?$select=${FOLDER_SELECT}`,
      );
      out[name] = { ...node, wellKnown: name };
    } catch (e) {
      if (statusOf(e) !== 404) throw e;
    }
  }
  return out;
}

/** Top-level folders, without the `searchfolders` root: search folders are
 *  saved queries over other folders, not places mail lives. */
export async function listTopFolders(client: GraphClient): Promise<MailFolderNode[]> {
  const [all, known] = await Promise.all([
    listPaged(client, `${GRAPH_BASE}/me/mailFolders?$top=100&$select=${FOLDER_SELECT}`),
    resolveWellKnown(client, ['searchfolders']),
  ]);
  const search = known.searchfolders?.id;
  return all.filter((n) => n.id !== search);
}

export function listChildFolders(client: GraphClient, id: string): Promise<MailFolderNode[]> {
  return listPaged(
    client,
    `${GRAPH_BASE}/me/mailFolders/${encodeURIComponent(id)}/childFolders?$top=100&$select=${FOLDER_SELECT}`,
  );
}

/** Every folder under the selected roots (roots included), mapped to the
 *  root that covers it. Each root's subtree is walked breadth-first, roots
 *  in order, so when roots overlap (a folder and its own descendant both
 *  selected) the first root in `rootIds` wins.
 *
 *  A folder whose own listing 404s was deleted upstream: that is a complete
 *  answer, not a partial one, so it is dropped (with `warn`) rather than
 *  failing discovery — otherwise one deleted root would fail every pull
 *  AND the Manage picker needed to untick it. Any other error propagates. */
export async function discoverTracked(
  client: GraphClient,
  rootIds: string[],
  warn: (msg: string) => void = () => {},
): Promise<Map<string, string>> {
  const tracked = new Map<string, string>();
  for (const root of rootIds) {
    if (tracked.has(root)) continue;
    tracked.set(root, root);
    const queue: Array<{ id: string; hasChildren: boolean }> = [{ id: root, hasChildren: true }];
    while (queue.length > 0) {
      const { id, hasChildren } = queue.shift()!;
      if (!hasChildren) continue;
      let children: MailFolderNode[];
      try {
        // eslint-disable-next-line no-await-in-loop
        children = await listChildFolders(client, id);
      } catch (e) {
        if (statusOf(e) !== 404) throw e;
        tracked.delete(id);
        warn(`ms365: mail folder ${id} no longer exists — skipped`);
        continue;
      }
      for (const child of children) {
        if (tracked.has(child.id)) continue;
        tracked.set(child.id, root);
        queue.push({ id: child.id, hasChildren: child.childFolderCount > 0 });
      }
    }
  }
  return tracked;
}
