/**
 * Microsoft 365 (Outlook mail) v2 source: platform-owned OAuth connect (the
 * `microsoft` OAuth provider — see manifest.json's `contributes.sources`
 * binding), a tenant-kind probe in place of legacy's id_token `tid`-claim
 * decode (v2 has no id_token to read), per-folder delta backfill + delta
 * sweep, and a pure `toDocument`.
 *
 * Ported from alpha-cent's v1 connector (`src/main/connectors/ms365/*.ts` +
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
} from './kiagent-contracts';
import { GraphClient, statusOf, type GraphClientDeps } from './graph-client';
import { GRAPH_BASE } from './graph-api';
import type { Ms365Cursor } from './cursor';
import { runBackfill } from './backfill';
import { runDelta } from './delta';
import { toDocument, type Ms365ThreadItem } from './to-document';

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
 * personal (MSA) accounts have no organization to enumerate and get a
 * 401/403/404 (or an empty `value`); a work/school tenant returns at least
 * one organization row. A 401 here specifically is NOT a reauth signal — see
 * `statusOf`'s doc comment for why it is safe to treat like 403/404 at this
 * call site.
 */
async function probeTenantKind(client: GraphClient): Promise<TenantKind> {
  try {
    const r = await client.request<{ value?: unknown[] }>(`${GRAPH_BASE}/organization`);
    return Array.isArray(r.value) && r.value.length > 0 ? 'work' : 'personal';
  } catch (e) {
    const status = statusOf(e);
    if (status === 401 || status === 403 || status === 404) return 'personal';
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

      return { identifier, config: { tenantKind } };
    },

    async *pull(session: Session, cursor: Ms365Cursor | null) {
      const client = clientFor(session);
      const tenantKind = tenantKindOf(session);
      if (cursor?.phase === 'live') {
        yield* runDelta(client, session, tenantKind, cursor);
      } else {
        yield* runBackfill(client, session, tenantKind, cursor);
      }
    },

    toDocument(item: Ms365ThreadItem): DocumentInput | null {
      return toDocument(item);
    },
  };
}
