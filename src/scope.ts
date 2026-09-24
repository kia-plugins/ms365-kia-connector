/**
 * Folder scope config (spec §3.1). `config.folderRoots` holds the selected
 * roots (Graph folder id + display name); each covers its subtree. An
 * account connected before folder scope has no `folderRoots` and keeps
 * today's behaviour until its first Save: Inbox + Sent Items enumerated,
 * no reconcile (core skips it for an undeclared scope).
 */
import type { FolderRootSelection } from '@kiagent/connector-sdk';
import type { GraphClient } from './graph-client';
import { resolveWellKnown } from './folders';

export const NEW_ACCOUNT_DEFAULTS = ['inbox', 'sentitems', 'archive'] as const;
export const LEGACY_ENUMERATION = ['inbox', 'sentitems'] as const;

/** The configured selection, or `null` for a legacy account. */
export function configuredRoots(config: Record<string, unknown>): FolderRootSelection[] | null {
  const { folderRoots } = config;
  if (!Array.isArray(folderRoots)) return null;
  return folderRoots.filter(
    (r): r is FolderRootSelection =>
      typeof r === 'object' &&
      r !== null &&
      typeof (r as FolderRootSelection).id === 'string' &&
      typeof (r as FolderRootSelection).name === 'string',
  );
}

/** Well-known folder names → roots, in the given order; a folder the
 *  mailbox lacks (no Archive) is omitted. */
export async function wellKnownRoots(
  client: GraphClient,
  names: readonly string[],
): Promise<FolderRootSelection[]> {
  const found = await resolveWellKnown(client, [...names]);
  return names
    .filter((n) => found[n])
    .map((n) => ({ id: found[n].id, name: found[n].displayName }));
}

/** The roots a pull enumerates: the configured selection, or for a legacy
 *  account the well-known Inbox + Sent Items. */
export async function effectiveRoots(
  client: GraphClient,
  config: Record<string, unknown>,
): Promise<{ roots: FolderRootSelection[]; legacy: boolean }> {
  const roots = configuredRoots(config);
  if (roots !== null) return { roots, legacy: false };
  return { roots: await wellKnownRoots(client, LEGACY_ENUMERATION), legacy: true };
}
