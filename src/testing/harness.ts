/**
 * Shared offline test harness: scripted host-shaped fetch responses (status /
 * statusText / lowercase headers / body: Uint8Array — see src/graph-client.ts),
 * a fake Graph world router, and fakes for Session / AuthChannel. No network,
 * no timers (client sleep/random are injected as instant/zero by the tests).
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
import type { NetFetch } from '../graph-client';
import type { GraphMessage } from '../parser';

export interface HostResponse {
  status: number;
  statusText: string;
  headers: Record<string, string>;
  body: Uint8Array;
}

export const jsonRes = (
  status: number,
  body: unknown,
  headers: Record<string, string> = {},
): HostResponse => ({
  status,
  statusText: '',
  headers,
  body: new TextEncoder().encode(JSON.stringify(body)),
});

const isHostResponse = (v: unknown): v is HostResponse =>
  typeof v === 'object' &&
  v !== null &&
  typeof (v as HostResponse).status === 'number' &&
  (v as HostResponse).body instanceof Uint8Array;

/** Instant clock + zero jitter for the source/client test seam. */
export const instantClock = { sleep: async () => {}, random: () => 0 };

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
  junkFolderId?: string;
  trashFolderId?: string;
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
  /** Checked first for every request; return undefined to fall through to
   *  the tables above. `count` is the per-exact-URL call number (0-based) —
   *  handy for "fails N times then succeeds" retry fixtures. */
  custom?: (url: URL, count: number) => HostResponse | undefined;
}

export function graphFetch(world: GraphWorld = {}): {
  fetchFn: NetFetch;
  calls: string[];
} {
  const calls: string[] = [];
  const counts = new Map<string, number>();
  const fetchFn: NetFetch = async (rawUrl) => {
    const urlStr = String(rawUrl);
    calls.push(urlStr);
    const count = counts.get(urlStr) ?? 0;
    counts.set(urlStr, count + 1);
    const url = new URL(urlStr);

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
    if (p === '/v1.0/me/mailFolders/junkemail') {
      return jsonRes(200, { id: world.junkFolderId ?? 'JUNK' });
    }
    if (p === '/v1.0/me/mailFolders/deleteditems') {
      return jsonRes(200, { id: world.trashFolderId ?? 'TRASH' });
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
    const v = world.urls?.[urlStr];
    if (v !== undefined) return isHostResponse(v) ? v : jsonRes(200, v);
    throw new Error(`fake graph: unhandled url ${urlStr}`);
  };
  return { fetchFn, calls };
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
