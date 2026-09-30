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
  Account,
  AuthChannel,
  Credentials,
  Document,
  DocumentInput,
  ExternalRef,
  FolderNode,
  FolderScopeUpdate,
  FolderSelectionChannel,
  HostFor,
  Session,
  Source,
} from '@kiagent/connector-sdk';
import { GraphClient, statusOf, type GraphClientDeps } from './graph-client';
import { downloadAttachment, GRAPH_BASE } from './graph-api';
import {
  rescope,
  type LegacyMs365Cursor,
  type Ms365Cursor,
} from './cursor';
import {
  ancestorsOf,
  discoverTracked,
  listChildFolders,
  listTopFolders,
  resolveWellKnown,
  type MailFolderNode,
} from './folders';
import { sync } from './sync';
import { EMAIL_THREAD_DOCUMENT_TYPE, toDocument, type Ms365ThreadItem } from './to-document';
import { configuredRoots, NEW_ACCOUNT_DEFAULTS, resolveScope, wellKnownRoots } from './scope';
import { listConversationIds } from './membership';
import { calendarToDocument } from './calendar/document';
import { listCalendars } from './calendar/graph';
import { calendarRefs } from './calendar/reconcile';
import { pullCalendars, selectedCalendars } from './calendar/pull';
import type { CalendarItem, GraphCalendar } from './calendar/types';

/**
 * Graph resource scopes only — legacy's SCOPES (`openid email profile
 * offline_access Mail.Read User.Read`) minus every OIDC/id_token-only scope
 * (`openid`, `email`, `profile`) AND `offline_access`: v2 has no id_token to
 * consume (tenant kind and identity both come from live Graph calls instead
 * — see `connect` below) and the platform's `microsoft` OAuth provider owns
 * refresh-token issuance/rotation on its own, outside any scope this source
 * requests. `Calendars.Read` since 3.0.0: the calendar half (src/calendar).
 */
export const SCOPES = ['Mail.Read', 'Calendars.Read', 'User.Read'];

/** What a pull yields: mail threads, then calendar occurrences. */
export type Ms365Item = Ms365ThreadItem | CalendarItem;

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

/** The stored cursor, or null to enumerate from scratch. Any cursor written
 *  before attachment children (and immutable ids) — a v1 cursor included —
 *  restarts once so every conversation re-emits with its attachments. */
function loadCursor(stored: unknown): Ms365Cursor | null {
  const c = stored as Ms365Cursor | LegacyMs365Cursor | null;
  return c !== null && 'v' in c && c.attachments === 1 ? c : null;
}

const JUNK_SUFFIX = ' (may contain phishing)';
const MANAGE_NOTE = 'Mail outside the selected folders will be removed from the index';

/** The scope Save's archive set (spec §3.4): conversations with a message in
 *  a folder that leaves the tracked set and none in the new tracked set.
 *  Held in memory only for the leaving folders; the (larger) staying
 *  listing streams against it. */
async function leavingRefs(
  client: GraphClient,
  prior: ReadonlyMap<string, string>,
  next: ReadonlyMap<string, string>,
): Promise<ExternalRef[]> {
  const leavingFolders = [...prior.keys()].filter((id) => !next.has(id));
  if (leavingFolders.length === 0) return []; // pure widening: no listing
  const leaving = new Set<string>();
  for await (const ids of listConversationIds(client, leavingFolders)) {
    for (const id of ids) leaving.add(id);
  }
  for await (const ids of listConversationIds(client, next.keys())) {
    for (const id of ids) leaving.delete(id);
    if (leaving.size === 0) break;
  }
  return [...leaving].map((externalId) => ({ externalId, type: EMAIL_THREAD_DOCUMENT_TYPE }));
}

export function createMs365Source(
  host: HostFor<'net'>,
  // Test seam only: GraphClient's sleep/random are injectable so retry tests
  // never actually wait; production callers omit this.
  clock?: Pick<GraphClientDeps, 'sleep' | 'random'>,
): Source<Ms365Cursor, Ms365Item> {
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
      documentTypes: ['email.thread', 'attachment', 'calendar.event'],
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
      const migrated = loadCursor(cursor);
      const start = rescope(
        migrated ?? {
          v: 2,
          phase: 'enumerate',
          folders: {},
          pending: [],
          retry: [],
          attachments: 1,
        },
        scope.tracked,
      );
      const tenantKind = tenantKindOf(session);
      const priorCal = migrated?.calendar;
      let last: Ms365Cursor = { ...start, calendar: priorCal };
      let phase: 'backfill' | 'live' = 'live';
      for await (const b of sync(client, session, tenantKind, scope, start)) {
        // Batch cursors replace wholesale: every mail batch carries the
        // calendar half unchanged.
        last = { ...b.cursor, calendar: priorCal };
        phase = b.phase;
        yield { ...b, cursor: last };
      }
      // Paused or quitting: sync() stopped early, and so does the calendar.
      if (session.signal.aborted) return;
      let cal;
      try {
        cal = await pullCalendars(
          client,
          session.account.config ?? {},
          priorCal,
          Date.now(),
          tenantKind,
        );
      } catch (e) {
        // No consent yet (an account from before 3.0.0): mail keeps syncing
        // until the user reconnects (spec §6).
        if (statusOf(e) === 403 && !priorCal) {
          session.log('warn', 'ms365: calendar skipped — reconnect the account to grant Calendars.Read');
          return;
        }
        throw e;
      }
      // Item batches carry the OLD calendar cursor and the last batch
      // commits the new one with the deletions; a calendar with no events
      // still needs that one batch to fix `since`.
      for (let i = 0; i < cal.items.length; i += 100) {
        yield { phase, items: cal.items.slice(i, i + 100), cursor: last };
      }
      yield { phase, items: [], deletions: cal.deletions, cursor: { ...last, calendar: cal.cursor } };
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
        onRequest: () => (requests += 1),
      })) {
        yield ids.map((externalId) => ({ externalId, type: EMAIL_THREAD_DOCUMENT_TYPE }));
      }
      // Core archives every unlisted document of the account, of any type:
      // the calendar events are listed too (spec §6).
      const calRefs = await calendarRefs(
        client,
        config,
        (session.account.cursor as Ms365Cursor | null)?.calendar,
        Date.now(),
      );
      if (calRefs.length) yield calRefs;
      session.log(
        'info',
        `ms365 reconcile: ${tracked.size} folders listed in ${requests} requests`,
      );
    },

    /** Spec §3.4: the Tracked folders picker over the Outlook folder tree.
     *  Persists nothing — core applies the returned config, cursor and
     *  `archiveRefs` in one transaction. */
    /** Reconnect: a fresh sign-in with the current scopes (how an account
     *  from before 3.0.0 grants Calendars.Read), refused for a different
     *  Microsoft account. */
    async reauthenticate(account: Account, auth: AuthChannel) {
      auth.status('Waiting for Microsoft sign-in…');
      const creds = await auth.oauth(SCOPES);
      const token = creds.accessToken;
      if (!token) throw new Error('ms365: Microsoft sign-in returned no access token');
      auth.status('Verifying the Microsoft account…');
      const client = new GraphClient({ fetch: host.net.fetch, getToken: async () => token, ...clock });
      const me = await client.request<{ mail?: string | null; userPrincipalName?: string }>(
        `${GRAPH_BASE}/me?$select=mail,userPrincipalName`,
      );
      const who = me.mail || me.userPrincipalName || '';
      const fold = (x: string) => x.trim().toLowerCase();
      if (fold(who) !== fold(account.identifier)) {
        throw new Error(
          `ms365: signed in as ${who}, but this account is ${account.identifier} — sign in with the original Microsoft account`,
        );
      }
    },

    async manageFolders(
      session: Session,
      channel: FolderSelectionChannel,
    ): Promise<FolderScopeUpdate<Ms365Cursor>> {
      const client = clientFor(session);
      const config = session.account.config ?? {};
      const warn = (m: string) => session.log('warn', m);
      const prior = await resolveScope(client, config, warn);
      const known = await resolveWellKnown(client, ['junkemail', 'msgfolderroot']);
      const toNode = (f: MailFolderNode): FolderNode => ({
        id: f.id,
        name: f.id === known.junkemail?.id ? `${f.displayName}${JUNK_SUFFIX}` : f.displayName,
        hasChildren: f.childFolderCount > 0,
      });
      const priorIds = prior.roots.map((r) => r.id);
      const allCals = await listCalendars(client).catch((e) => {
        // No Calendars.Read consent yet: the Calendars tab stays empty.
        if (statusOf(e) === 403) return [] as GraphCalendar[];
        throw e;
      });
      const priorCals = selectedCalendars(config, allCals);
      const calIds = new Set(allCals.map((c) => c.id));
      const calNode = (c: { id: string; name: string }): FolderNode => ({
        id: c.id,
        name: c.name,
        hasChildren: false,
      });
      const picked = await channel.pickFolders({
        modes: [
          { key: 'mail', label: 'Mail folders' },
          { key: 'calendars', label: 'Calendars' },
        ],
        multiSelect: true,
        purpose: 'manage',
        note: MANAGE_NOTE,
        selected: [
          ...prior.roots.map((r) => ({
            id: r.id,
            name: r.name,
            hasChildren: [...prior.tracked].some(([f, root]) => root === r.id && f !== r.id),
          })),
          ...priorCals.map(calNode),
        ],
        expand: await ancestorsOf(client, priorIds, known.msgfolderroot?.id),
        roots: async (mode) =>
          mode === 'calendars'
            ? allCals.map(calNode)
            : (await listTopFolders(client)).map(toNode),
        children: async (id) =>
          calIds.has(id) ? [] : (await listChildFolders(client, id)).map(toNode),
      });
      const pickedCals = picked.filter((n) => calIds.has(n.id));
      const pickedMail = picked.filter((n) => !calIds.has(n.id));
      if (pickedMail.length === 0) throw new Error('ms365: no mail folders selected');

      // Retained roots in prior order, then new ones in pick order.
      const pickedIds = new Set(pickedMail.map((n) => n.id));
      const folderRoots = [
        ...prior.roots.filter((r) => pickedIds.has(r.id)),
        ...pickedMail
          .filter((n) => !priorIds.includes(n.id))
          .map((n) => ({
            id: n.id,
            name: n.name.endsWith(JUNK_SUFFIX) ? n.name.slice(0, -JUNK_SUFFIX.length) : n.name,
          })),
      ];
      const next = await discoverTracked(
        client,
        folderRoots.map((r) => r.id),
        warn,
      );
      // A legacy account's first Save lists nothing: core grants the first
      // declaration a reconcile allowance, and that pass archives exactly
      // indexed − staying (spec §3.4).
      const archiveRefs = prior.legacy ? [] : await leavingRefs(client, prior.tracked, next);

      // Calendars: an explicit list from now on ([] = calendar off), but only
      // when the tab listed any: without consent (403) or calendars, a mail
      // Save must not switch the calendar off. An unticked calendar's events
      // carry its id as scopeRootId, so archiveScopeRootIds archives them.
      const calendarRoots = pickedCals.map((n) => ({ id: n.id, name: n.name }));
      const keepCal = new Set(calendarRoots.map((r) => r.id));
      const calConfig = allCals.length ? { calendarRoots } : {};
      const migrated = loadCursor(session.account.cursor);
      const cursor = migrated && rescope(migrated, next);
      return {
        config: { ...config, folderRoots, ...calConfig },
        cursor:
          cursor && cursor.calendar && allCals.length
            ? {
                ...cursor,
                calendar: {
                  ...cursor.calendar,
                  cals: Object.fromEntries(
                    Object.entries(cursor.calendar.cals).filter(([id]) => keepCal.has(id)),
                  ),
                },
              }
            : cursor,
        archiveScopeRootIds: priorCals.map((c) => c.id).filter((id) => !keepCal.has(id)),
        reattributeScopeRoots: [],
        archiveRefs,
      };
    },

    toDocument(item: Ms365Item): DocumentInput | DocumentInput[] | null {
      if ('calendarEvent' in item) return calendarToDocument(item);
      return toDocument(item);
    },

    /** Bytes of one attachment child, re-resolved by name + size under its
     *  immutable message id (to-document.ts). `null` = gone upstream; a
     *  throw (auth, network) is transient — core defers and retries. */
    async fetchBytes(session: Session, doc: Document): Promise<Uint8Array | null> {
      if (doc.type !== 'attachment') return null;
      const meta = doc.metadata as { messageId?: unknown; filename?: unknown; sizeBytes?: unknown };
      if (typeof meta.messageId !== 'string') return null;
      return downloadAttachment(
        clientFor(session),
        meta.messageId,
        typeof meta.filename === 'string' ? meta.filename : null,
        typeof meta.sizeBytes === 'number' ? meta.sizeBytes : null,
      );
    },
  };
}
