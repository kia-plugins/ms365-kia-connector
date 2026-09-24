/**
 * Microsoft 365 (Outlook mail) v2 source: platform-owned OAuth connect (the
 * `microsoft` OAuth provider — see manifest.json's `contributes.sources`
 * binding), a tenant-kind probe in place of legacy's id_token `tid`-claim
 * decode (v2 has no id_token to read), a pull over the selected folder tree
 * (sync.ts), and a pure `toDocument`.
 *
 * Ported from the v1 connector (`src/main/connectors/ms365/*.ts` +
 * `ms-shared/graph-fetch.ts` + `ms-shared/walk-delta.ts`): `client.ts`'s
 * `resolveExcludedFolderIds` (junk/deleted folders excluded from every
 * sweep), `backfill.ts`'s two-phase enumerate/ingest, `delta.ts`'s
 * per-folder deltaLink walk + 410 syncStateNotFound re-prime, `parser.ts`
 * verbatim, and `thread-builder.ts`'s markdown shape (minus the DB/attachment
 * plumbing — see to-document.ts and graph-api.ts for what changed and why).
 *
 * NOT ported (platform-owned in v2, so there is nothing left for this
 * source to do):
 *  - PKCE / authorize-url / code-exchange / refresh (`ms-shared/oauth.ts`) —
 *    the platform's `microsoft` OAuth provider owns the whole flow;
 *    `auth.oauth(SCOPES)` is the only touchpoint.
 *  - id_token decoding (`decodeTenantIdFromIdToken`/`resolveTenantKind`) —
 *    v2's `auth.oauth()` never returns an id_token (see contracts.ts's
 *    `Credentials` — no such field). This port derives the SAME `TenantKind`
 *    result via a live probe instead (`probeTenantKind` below).
 *  - The Azure `clientId`/`tenantId` account-schema fields and
 *    `client-credentials.ts` — the platform holds all OAuth client
 *    credentials; this connector holds none (see README's Privacy section).
 */
import type {
  AuthChannel,
  Credentials,
  DocumentInput,
  HostFor,
  Session,
  Source,
} from '@kiagent/connector-sdk';
import { GraphClient, statusOf, type GraphClientDeps } from './graph-client';
import { GRAPH_BASE } from './graph-api';
import {
  migrateCursor,
  rescope,
  type LegacyMs365Cursor,
  type Ms365Cursor,
} from './cursor';
import { resolveWellKnown } from './folders';
import { sync } from './sync';
import { EMAIL_THREAD_DOCUMENT_TYPE, toDocument, type Ms365ThreadItem } from './to-document';
import { configuredRoots, NEW_ACCOUNT_DEFAULTS, resolveScope, wellKnownRoots } from './scope';
import { listConversationIds } from './membership';

/**
 * Graph resource scopes only — legacy's SCOPES (`openid email profile
 * offline_access Mail.Read User.Read`) minus every OIDC/id_token-only scope
 * (`openid`, `email`, `profile`) AND `offline_access`: v2 has no id_token to
 * consume (tenant kind and identity both come from live Graph calls instead
 * — see `connect` below) and the platform's `microsoft` OAuth provider owns
 * refresh-token issuance/rotation on its own, outside any scope this source
 * requests.
 */
export const SCOPES = ['Mail.Read', 'User.Read'];

type TenantKind = 'work' | 'personal';

function tenantKindOf(session: Session): TenantKind {
  const cfg = session.account.config as { tenantKind?: unknown };
  return cfg.tenantKind === 'work' ? 'work' : 'personal';
}

/**
 * `GET /organization` in place of legacy's id_token `tid`-claim decode:
 * personal (MSA) accounts have no organization to enumerate and fail the call
 * (or return an empty `value`); a work/school tenant returns at least one
 * organization row. A 401 here specifically is NOT a reauth signal — see
 * `statusOf`'s doc comment for why it is safe to treat like 403/404 at this
 * call site.
 *
 * Every non-retryable 4xx counts as personal, not just 401/403/404: MSA
 * accounts answer 400 BadRequest — "This API is not supported for MSA accounts
 * (no addressUrl for Microsoft.DirectoryServices,False)" — which an
 * enumerated allow-list missed, failing connect outright for every personal
 * account. This is a heuristic probe with a safe fallback, so the only
 * failures worth propagating are the ones a retry could fix (429/5xx, per
 * `isRetryableGraphFailure`).
 */
async function probeTenantKind(client: GraphClient): Promise<TenantKind> {
  try {
    const r = await client.request<{ value?: unknown[] }>(`${GRAPH_BASE}/organization`);
    return Array.isArray(r.value) && r.value.length > 0 ? 'work' : 'personal';
  } catch (e) {
    const status = statusOf(e);
    if (status !== undefined && status >= 400 && status < 500 && status !== 429) {
      return 'personal';
    }
    throw e;
  }
}

async function requireToken(session: Session): Promise<string> {
  const creds = await session.credentials();
  if (!creds?.accessToken) {
    throw new Error('ms365: no credentials available — reconnect the account');
  }
  return creds.accessToken;
}

/** The ids a v1 cursor's `inbox`/`sentitems` keys stood for. */
async function legacyFolderIds(client: GraphClient): Promise<{ inbox: string; sentitems: string }> {
  const known = await resolveWellKnown(client, ['inbox', 'sentitems']);
  if (!known.inbox || !known.sentitems) {
    throw new Error('ms365: cannot resolve Inbox / Sent Items to migrate the sync cursor');
  }
  return { inbox: known.inbox.id, sentitems: known.sentitems.id };
}

export function createMs365Source(
  host: HostFor<'net'>,
  // Test seam only: GraphClient's sleep/random are injectable so retry tests
  // never actually wait; production callers omit this.
  clock?: Pick<GraphClientDeps, 'sleep' | 'random'>,
): Source<Ms365Cursor, Ms365ThreadItem> {
  const clientFor = (session: Session): GraphClient =>
    new GraphClient({
      fetch: host.net.fetch,
      getToken: () => requireToken(session),
      ...clock,
    });

  return {
    descriptor: {
      id: 'ms365',
      name: 'Microsoft 365',
      documentTypes: ['email.thread'],
      auth: 'oauth',
      multiAccount: true,
      cadence: { every: '15m' },
      folderScope: true,
    },

    async connect(auth: AuthChannel) {
      auth.status('Waiting for Microsoft sign-in…');
      const creds: Credentials = await auth.oauth(SCOPES);
      const accessToken = creds.accessToken;
      if (!accessToken) {
        throw new Error('ms365: Microsoft sign-in returned no access token');
      }
      const client = new GraphClient({
        fetch: host.net.fetch,
        getToken: async () => accessToken,
        ...clock,
      });

      auth.status('Fetching Microsoft 365 profile…');
      const me = await client.request<{ mail?: string | null; userPrincipalName?: string }>(
        `${GRAPH_BASE}/me?$select=mail,userPrincipalName`,
      );
      const identifier = me.mail || me.userPrincipalName;
      if (!identifier) {
        throw new Error('ms365: /me response missing both mail and userPrincipalName');
      }

      auth.status('Checking Microsoft 365 account type…');
      const tenantKind = await probeTenantKind(client);
      const folderRoots = await wellKnownRoots(client, NEW_ACCOUNT_DEFAULTS);

      return { identifier, config: { tenantKind, folderRoots } };
    },

    async *pull(session: Session, cursor: Ms365Cursor | null) {
      const client = clientFor(session);
      const scope = await resolveScope(client, session.account.config ?? {}, (m) =>
        session.log('warn', m),
      );
      const stored = cursor as Ms365Cursor | LegacyMs365Cursor | null;
      const migrated = migrateCursor(
        stored,
        stored && !('v' in stored) ? await legacyFolderIds(client) : { inbox: '', sentitems: '' },
      );
      const start = rescope(
        migrated ?? { v: 2, phase: 'enumerate', folders: {}, pending: [], retry: [] },
        scope.tracked,
      );
      yield* sync(client, session, tenantKindOf(session), scope, start);
    },

    /** Spec §3.3: every conversation with a message in the tracked tree.
     *  Core never calls this for an account without a declared scope
     *  (legacy); the throw only guards that contract. A discovery failure
     *  rejects before anything is yielded. */
    async *reconcile(session: Session) {
      const config = session.account.config ?? {};
      if (configuredRoots(config) === null) {
        throw new Error('ms365: reconcile without declared scope');
      }
      const client = clientFor(session);
      const warn = (m: string) => session.log('warn', m);
      const { tracked } = await resolveScope(client, config, warn);
      let requests = 0;
      for await (const ids of listConversationIds(client, tracked.keys(), {
        signal: session.signal,
        warn,
        onRequest: () => (requests += 1),
      })) {
        yield ids.map((externalId) => ({ externalId, type: EMAIL_THREAD_DOCUMENT_TYPE }));
      }
      session.log(
        'info',
        `ms365 reconcile: ${tracked.size} folders listed in ${requests} requests`,
      );
    },

    toDocument(item: Ms365ThreadItem): DocumentInput | null {
      return toDocument(item);
    },
  };
}
