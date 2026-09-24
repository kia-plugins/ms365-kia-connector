/**
 * Shared offline test harness: what is left of it after the generic half moved
 * to `@kiagent/connector-sdk/testing` (which was itself generalized FROM this
 * file). The kit now owns host-shaped responses (`jsonRes`), the exact-URL
 * scripted-fetch skeleton with its per-URL call counts (`scriptedFetch`), and
 * the zero-wait clock (`instantClock`); what stays here is Graph-specific: the
 * `GraphWorld` fixture shape, its path router, and the Session / AuthChannel /
 * host fakes carrying this connector's own defaults.
 *
 * No network, no timers (client sleep/random are injected as instant/zero by
 * the tests).
 *
 * Lives outside src/__tests__ so jest's default testMatch does not treat it
 * as a suite. Never bundled: build.mjs only follows imports from index.ts.
 */
import type {
  Account,
  AuthChannel,
  Credentials,
  HostFor,
  Session,
} from '@kiagent/connector-sdk';
import type { HostResponse } from '@kiagent/connector-sdk/http';
import { jsonRes, scriptedFetch } from '@kiagent/connector-sdk/testing';
import type { NetFetch } from '../graph-client';
import type { MailFolderNode } from '../folders';
import type { GraphMessage } from '../parser';

/** Re-exported so this harness stays the single import site for the tests:
 *  the generic pieces now come from the SDK kit, the Graph-specific ones
 *  below are still local. */
export { jsonRes, instantClock } from '@kiagent/connector-sdk/testing';
export type { HostResponse } from '@kiagent/connector-sdk/http';

interface ConversationPageFx {
  value?: GraphMessage[];
  '@odata.nextLink'?: string;
}

/** The fake upstream Microsoft Graph. Every field optional; unhandled
 *  requests throw loudly (the client treats that as a network error and
 *  retries — tests run with instant sleep, so a genuinely missing fixture
 *  still fails fast). */
export interface GraphWorld {
  about?: { mail?: string | null; userPrincipalName?: string };
  /** `/organization` probe response. Omit for the default ('personal':
   *  200 with an empty `value` array). */
  organization?: { status: number; body?: unknown } | { value: unknown[] };
  /** Exact-URL-keyed response table — covers folder delta pages, whose
   *  nextLink/deltaLink are literal strings the fixture itself defines
   *  (mirroring the legacy nock-based tests, which registered exact
   *  next/delta-link URLs the same way). Values are JSON bodies (200) unless
   *  already a HostResponse (for non-200 / custom-header cases). */
  urls?: Record<string, HostResponse | unknown>;
  /** conversationId -> full message list (unpaginated) OR explicit pages
   *  (array of arrays) for the `/me/messages?$filter=conversationId eq …`
   *  endpoint. A conversationId absent from this table AND from `custom`
   *  throws (fails the test loudly) rather than silently returning empty —
   *  an explicit `{ [id]: [] }` entry is how a test represents "deleted /
   *  zero-message conversation". */
  conversations?: Record<string, GraphMessage[] | GraphMessage[][]>;
  /** The mail folder tree. `top` and each `children` list are single
   *  pages unless given as explicit pages (array of arrays), paged with a
   *  `pageToken` nextLink like `conversations`. A well-known name absent
   *  from `wellKnown` answers 404. Omitted → `DEFAULT_FOLDERS`. */
  folders?: {
    top: MailFolderNode[] | MailFolderNode[][];
    children?: Record<string, MailFolderNode[] | MailFolderNode[][]>;
    wellKnown?: Record<string, MailFolderNode>;
  };
  /** folderId → the conversationIds of its messages, for the reconcile /
   *  manageFolders listing `/me/mailFolders/{id}/messages?$select=
   *  conversationId` (one page, or explicit pages). A folder absent here
   *  throws (fails the test loudly). */
  folderMessages?: Record<string, string[] | string[][]>;
  /** Checked first for every request; return undefined to fall through to
   *  the tables above. `count` is the per-exact-URL call number (0-based) —
   *  handy for "fails N times then succeeds" retry fixtures. */
  custom?: (url: URL, count: number) => HostResponse | undefined;
}

export function graphFetch(world: GraphWorld = {}): {
  fetchFn: NetFetch;
  calls: string[];
} {
  /** The Graph-domain router, layered onto the SDK kit's `scriptedFetch` as
   *  its `custom` callback — which the kit consults BEFORE its exact-URL
   *  table, so the fall-through order is the original's exactly:
   *  `world.custom` → these Graph paths → `world.urls` → "unhandled url".
   *
   *  Returning `undefined` on no match is load-bearing: throwing here would
   *  short-circuit the kit before it ever reads `world.urls` (which is what
   *  the delta/backfill fixtures are built from). The two in-path throws
   *  below are deliberate — they fire only once a request HAS matched
   *  `/me/messages` but carries no fixture. */
  const route = (url: URL, count: number): HostResponse | undefined => {
    if (world.custom) {
      const r = world.custom(url, count);
      if (r) return r;
    }

    const p = url.pathname;
    if (p === '/v1.0/me' && url.searchParams.get('$select') === 'mail,userPrincipalName') {
      return jsonRes(200, world.about ?? { mail: 'user@example.com' });
    }
    if (p === '/v1.0/organization') {
      const org = world.organization;
      if (!org) return jsonRes(200, { value: [] });
      if ('status' in org) return jsonRes(org.status, org.body ?? {});
      return jsonRes(200, { value: org.value });
    }
    const folders = world.folders ?? DEFAULT_FOLDERS;
    if (p === '/v1.0/me/mailFolders') {
      return jsonRes(200, pageOf(url, folders.top));
    }
    const kids = /^\/v1\.0\/me\/mailFolders\/([^/]+)\/childFolders$/.exec(p);
    if (kids) {
      const list = folders.children?.[decodeURIComponent(kids[1])];
      if (list === undefined) throw new Error(`fake graph: no children fixture for ${kids[1]}`);
      return jsonRes(200, pageOf(url, list));
    }
    const listed = /^\/v1\.0\/me\/mailFolders\/([^/]+)\/messages$/.exec(p);
    if (listed) {
      const ids = world.folderMessages?.[decodeURIComponent(listed[1])];
      if (ids === undefined) throw new Error(`fake graph: no folderMessages fixture for ${listed[1]}`);
      const page = pageOf(url, ids);
      return jsonRes(200, { ...page, value: page.value.map((conversationId) => ({ conversationId })) });
    }
    const named = /^\/v1\.0\/me\/mailFolders\/([^/]+)$/.exec(p);
    if (named) {
      const hit = folders.wellKnown?.[decodeURIComponent(named[1])];
      return hit
        ? jsonRes(200, hit)
        : jsonRes(404, { error: { code: 'ErrorFolderNotFound' } });
    }
    if (p === '/v1.0/me/messages') {
      const filter = url.searchParams.get('$filter') ?? '';
      const m = /conversationId eq '([^']+)'/.exec(filter);
      const conversationId = m?.[1];
      if (!conversationId) throw new Error(`fake graph: unparseable $filter ${filter}`);
      const raw = world.conversations?.[conversationId];
      if (raw === undefined) {
        throw new Error(`fake graph: no conversation fixture for ${conversationId}`);
      }
      const pages: GraphMessage[][] = Array.isArray(raw[0])
        ? (raw as GraphMessage[][])
        : raw.length
          ? [raw as GraphMessage[]]
          : [[]];
      const tok = url.searchParams.get('pageToken');
      const idx = tok ? Number(tok) : 0;
      const body: ConversationPageFx = { value: pages[idx] ?? [] };
      if (idx + 1 < pages.length) {
        const next = new URL(url.toString());
        next.searchParams.set('pageToken', String(idx + 1));
        body['@odata.nextLink'] = next.toString();
      }
      return jsonRes(200, body);
    }
    return undefined;
  };

  const { fetchFn, calls } = scriptedFetch({ urls: world.urls, custom: route });
  return { fetchFn, calls };
}

/** The mailbox every test gets unless it sets `folders`: the three
 *  well-known folders a new account's default selection resolves. */
export const DEFAULT_FOLDERS: NonNullable<GraphWorld['folders']> = {
  top: [],
  children: { 'INBOX-ID': [], 'SENT-ID': [], 'ARCHIVE-ID': [] },
  wellKnown: {
    inbox: { id: 'INBOX-ID', displayName: 'Inbox', childFolderCount: 0 },
    sentitems: { id: 'SENT-ID', displayName: 'Sent Items', childFolderCount: 0 },
    archive: { id: 'ARCHIVE-ID', displayName: 'Archive', childFolderCount: 0 },
  },
};

/** One page of a (possibly paged) fixture list, with a `pageToken`
 *  nextLink while pages remain. */
function pageOf<T>(url: URL, raw: T[] | T[][]): { value: T[]; '@odata.nextLink'?: string } {
  const pages: T[][] = raw.length && Array.isArray(raw[0]) ? (raw as T[][]) : [raw as T[]];
  const tok = url.searchParams.get('pageToken');
  const idx = tok ? Number(tok) : 0;
  const body: { value: T[]; '@odata.nextLink'?: string } = { value: pages[idx] ?? [] };
  if (idx + 1 < pages.length) {
    const next = new URL(url.toString());
    next.searchParams.set('pageToken', String(idx + 1));
    body['@odata.nextLink'] = next.toString();
  }
  return body;
}

export function makeHost(fetchFn: NetFetch): HostFor<'net'> {
  return {
    self: { id: 'kia.ms365', dataDir: '/tmp' },
    log: () => {},
    net: { fetch: fetchFn },
  };
}

export function makeSession(
  opts: {
    creds?: Credentials | null;
    config?: Record<string, unknown>;
    signal?: AbortSignal;
  } = {},
): { session: Session; logs: { level: string; msg: string }[] } {
  const logs: { level: string; msg: string }[] = [];
  const session: Session = {
    account: {
      id: 'acc-1',
      source: 'ms365',
      identifier: 'user@example.com',
      config: opts.config ?? {},
      status: 'live',
      cursor: null,
      createdAt: '2026-01-01T00:00:00Z',
    } as Account,
    signal: opts.signal ?? new AbortController().signal,
    credentials: async () =>
      opts.creds === undefined ? { accessToken: 'ms-test-token-deadbeef' } : opts.creds,
    log: (level, msg) => logs.push({ level, msg }),
  };
  return { session, logs };
}

export function makeAuth(opts: { creds?: Credentials } = {}): {
  auth: AuthChannel;
  statuses: string[];
  getScopes: () => string[] | undefined;
} {
  const statuses: string[] = [];
  let scopes: string[] | undefined;
  const auth: AuthChannel = {
    oauth: async (s) => {
      scopes = s;
      return opts.creds ?? { accessToken: 'ms-test-token-deadbeef' };
    },
    showQr: () => {},
    prompt: async () => ({}),
    pickFolders: async () => {
      throw new Error('ms365 does not use pickFolders');
    },
    status: (m) => statuses.push(m),
  };
  return { auth, statuses, getScopes: () => scopes };
}

export async function collect<T>(iter: AsyncIterable<T>): Promise<T[]> {
  const out: T[] = [];
  for await (const v of iter) out.push(v);
  return out;
}

/** Builds a raw Graph message fixture with sensible defaults, matching the
 *  legacy nock test bodies' shape. */
export function graphMsg(over: Partial<GraphMessage> = {}): GraphMessage {
  return {
    id: 'm1',
    conversationId: 'C1',
    internetMessageId: '<m1@x>',
    subject: 'hello',
    from: { emailAddress: { address: 'a@x.com', name: 'A' } },
    toRecipients: [],
    ccRecipients: [],
    receivedDateTime: '2026-05-20T10:00:00Z',
    body: { contentType: 'text', content: 'hi' },
    internetMessageHeaders: [],
    hasAttachments: false,
    parentFolderId: 'inbox',
    ...over,
  };
}
