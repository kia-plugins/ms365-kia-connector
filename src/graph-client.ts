/**
 * v2 port of the v1 ms365 Graph client (v1 repo
 * `src/main/connectors/ms-shared/graph-fetch.ts` +
 * `src/main/connectors/http-shared/bearer-fetch.ts`), reshaped to the
 * google-docs-kia-connector `DriveClient` idiom: a class over the host's
 * `net.fetch` surface rather than a positional-fetchImpl function.
 *
 * Preserved verbatim from v1:
 *  - Retry policy: up to MAX_RETRIES=4 retries after the initial request,
 *    backoff `min(60000, 1000*2^attempt) + jitter*250`, `Retry-After`
 *    (seconds) honored when finite and > 0.
 *  - Retryable = 429 or >=500 ONLY — v1's graph-fetch.ts is deliberately
 *    status-only (Graph throttles via 429/503 and never uses Google's
 *    403-with-quota-reason pattern; do not add a body regex here).
 *  - Thrown message CONTRACT: `graph <status> <url> <body>` — the delta
 *    sync-state-expired check (`isSyncStateExpired`) matches against this
 *    exact format (v1 matched `/410.*syncStateNotFound/i` the same way).
 *  - Token fetched fresh per attempt via the `getToken` seam (`pull` passes
 *    `session.credentials()`; `connect` passes the accessToken from
 *    `auth.oauth`).
 *
 * Deltas from v1:
 *  1. All I/O goes through `deps.fetch` — the host's `net.fetch` surface —
 *     never global fetch. The host resolves to a plain object (status /
 *     statusText / headers with lowercase keys / body: Uint8Array), so
 *     responses are decoded manually and there is no `.ok`.
 *  2. The v1 90s per-attempt AbortController timeout is DROPPED:
 *     `host.net.fetch` owns the transport (platform-level retry/backoff and
 *     socket hygiene), so the connector no longer arms its own timers.
 *  3. HTTP 401 throws `Ms365AuthError` (message ends "— reconnect the
 *     account"), is NEVER retried, and always propagates — the engine flips
 *     the account to needsReauth on auth errors. v1 had no 401 special-case
 *     (401s simply were not retried by the generic bearer-fetch retry
 *     predicate, but nothing distinguished them from other propagated
 *     errors).
 *  4. `sleep`/`random` are injectable so tests never actually wait.
 */

// The host's own net.fetch type (PluginNet['fetch']), via the SDK.
import type { NetFetch } from '@kiagent/connector-sdk/http';
export type { NetFetch };

/** Graph's opt-in for ids that survive folder moves. */
export const IMMUTABLE_IDS = 'IdType="ImmutableId"';

export type ResponseType = 'json' | 'text' | 'bytes';

/** Max retries AFTER the initial request (v1 bearer-fetch MAX_ATTEMPTS). */
const MAX_RETRIES = 4;
/** Error bodies are truncated to this many chars in thrown messages. */
const BODY_SNIPPET_CHARS = 500;

/** The host `net.fetch` surface resolves to this shape — header keys are
 *  lowercase (built via Object.fromEntries(res.headers.entries())). */
interface HostResponse {
  status: number;
  statusText: string;
  headers: Record<string, string>;
  body: Uint8Array;
}

/** Non-2xx Graph response (except 401). Message format is load-bearing —
 *  see the module doc. */
export class GraphApiError extends Error {
  /** Graph's `error.code` (e.g. ErrorItemNotFound), when the body had one. */
  public readonly code: string | undefined;

  constructor(
    public readonly status: number,
    public readonly url: string,
    body: string,
    fullBody: string = body,
  ) {
    super(`graph ${status} ${url} ${body}`);
    this.name = 'GraphApiError';
    this.code = graphErrorCode(fullBody);
  }
}

function graphErrorCode(body: string): string | undefined {
  try {
    const code = (JSON.parse(body) as { error?: { code?: unknown } })?.error?.code;
    return typeof code === 'string' ? code : undefined;
  } catch {
    return undefined;
  }
}

/** HTTP 401 (or missing credentials). Never retried, always propagated —
 *  every later call would fail identically. */
export class Ms365AuthError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'Ms365AuthError';
  }
}

export const isAuthError = (e: unknown): e is Ms365AuthError =>
  e instanceof Ms365AuthError;

/** Uniform HTTP status extraction across the two thrown-error shapes above —
 *  `undefined` for anything else (network errors, aborts). Used by
 *  `probeTenantKind` (source.ts), which treats 401 on `/organization`
 *  specially: by the time that probe runs, `connect()` has already made a
 *  successful `/me` call with the same token, so a 401 there is NOT a
 *  reauth signal — it means "this endpoint doesn't apply to a personal
 *  account", exactly like a 403/404 would. */
export function statusOf(e: unknown): number | undefined {
  if (e instanceof GraphApiError) return e.status;
  if (e instanceof Ms365AuthError) return 401;
  return undefined;
}

/** A 404 that means "this folder (or item) no longer exists" — NOT any 404:
 *  Graph also answers 404 for mailbox-level states (a mailbox not enabled
 *  for REST, `ResourceNotFound` during licensing or migration). Reading
 *  those as "folder deleted" would drop the whole tracked tree, so only
 *  the folder/item codes qualify; everything else propagates. */
export function isFolderGone(e: unknown): boolean {
  return (
    e instanceof GraphApiError &&
    e.status === 404 &&
    /"code"\s*:\s*"(ErrorItemNotFound|ErrorFolderNotFound)"/.test(e.message)
  );
}

/** 429 = throttled; 5xx = transient. Other 4xx are caller errors.
 *  Deliberately status-only (v1 graph-fetch.ts parity) — MS Graph never
 *  needs a body-regex quota check the way Google's Drive API does. */
export function isRetryableGraphFailure(status: number): boolean {
  return status === 429 || status >= 500;
}

/** Graph mail delta's 410 "the sync state is no longer valid" failure — v1
 *  matched `/410.*syncStateNotFound/i` against the whole bearer-fetch
 *  message; the GraphApiError message format preserves that same text. */
export function isSyncStateExpired(e: unknown): e is GraphApiError {
  return e instanceof GraphApiError && e.status === 410 && /syncStateNotFound/i.test(e.message);
}

export interface GraphClientDeps {
  fetch: NetFetch;
  /** Fresh token per attempt — see module doc. */
  getToken: () => Promise<string>;
  sleep?: (ms: number) => Promise<void>;
  /** Jitter source for backoff (default Math.random) — injectable so retry
   *  tests can assert exact delays. */
  random?: () => number;
}

export class GraphClient {
  private readonly fetchFn: NetFetch;

  private readonly getToken: () => Promise<string>;

  private readonly sleepFn: (ms: number) => Promise<void>;

  private readonly random: () => number;

  constructor(deps: GraphClientDeps) {
    this.fetchFn = deps.fetch;
    this.getToken = deps.getToken;
    this.sleepFn =
      deps.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.random = deps.random ?? Math.random;
  }

  async request<T>(
    url: string,
    opts: { responseType?: ResponseType; extraHeaders?: Record<string, string> } = {},
  ): Promise<T> {
    const responseType = opts.responseType ?? 'json';
    for (let attempt = 0; ; attempt++) {
      const token = await this.getToken(); // fresh per attempt
      let res: HostResponse | undefined;
      let netError: Error | undefined;
      try {
        res = (await this.fetchFn(url, {
          headers: {
            authorization: `Bearer ${token}`,
            accept: 'application/json',
            ...opts.extraHeaders,
            // Immutable ids on EVERY request: Outlook item ids otherwise
            // change when a message moves folders, and attachment children
            // are keyed (and hashed) by their message id.
            prefer: [IMMUTABLE_IDS, opts.extraHeaders?.prefer]
              .filter(Boolean)
              .join(', '),
          },
        })) as HostResponse;
      } catch (e) {
        netError = e instanceof Error ? e : new Error(String(e));
      }

      if (netError) {
        if (attempt < MAX_RETRIES) {
          await this.sleepFn(this.backoff(attempt));
          continue;
        }
        throw netError;
      }

      const r = res!;
      if (r.status >= 200 && r.status < 300) {
        if (responseType === 'bytes') return r.body as unknown as T;
        const text = new TextDecoder().decode(r.body);
        if (responseType === 'text') return text as unknown as T;
        return (text ? JSON.parse(text) : undefined) as T;
      }

      const body = new TextDecoder()
        .decode(r.body)
        .slice(0, BODY_SNIPPET_CHARS);
      if (r.status === 401) {
        throw new Ms365AuthError(
          `graph 401 ${url} ${body} — reconnect the account`,
        );
      }
      if (attempt < MAX_RETRIES && isRetryableGraphFailure(r.status)) {
        const retryAfterS = Number(r.headers['retry-after']);
        const delay =
          Number.isFinite(retryAfterS) && retryAfterS > 0
            ? retryAfterS * 1000
            : this.backoff(attempt);
        await this.sleepFn(delay);
        continue;
      }
      throw new GraphApiError(r.status, url, body);
    }
  }

  /**
   * One outbound POST (send a reply / a new mail). NOT `request()`'s retry
   * policy: a send is not idempotent, so a network error or 5xx — where
   * Graph may already have accepted the message — is thrown at once, never
   * retried. Only a 429 (rejected before processing) is retried, and only
   * while its wait still fits the ONE overall deadline, which also bounds
   * every attempt through the host fetch's own timeout: nothing this call
   * started may still be in flight after it has thrown.
   */
  async post(
    url: string,
    body: unknown,
    opts: { deadlineMs: number; now?: () => number },
  ): Promise<void> {
    const now = opts.now ?? Date.now;
    const deadline = now() + opts.deadlineMs;
    for (let attempt = 0; ; attempt++) {
      const token = await this.getToken();
      const remaining = deadline - now();
      if (remaining <= 0) throw new Error(`graph send timed out ${url}`);
      const r = (await this.fetchFn(url, {
        method: 'POST',
        headers: {
          authorization: `Bearer ${token}`,
          accept: 'application/json',
          'content-type': 'application/json',
          prefer: IMMUTABLE_IDS,
        },
        body: JSON.stringify(body),
        timeoutMs: Math.max(1, Math.floor(remaining)),
      } as never)) as HostResponse;
      if (r.status >= 200 && r.status < 300) return;
      const full = new TextDecoder().decode(r.body);
      const text = full.slice(0, BODY_SNIPPET_CHARS);
      if (r.status === 401) {
        throw new Ms365AuthError(`graph 401 ${url} ${text} — reconnect the account`);
      }
      if (r.status === 429 && attempt < MAX_RETRIES) {
        const retryAfterS = Number(r.headers['retry-after']);
        const delay =
          Number.isFinite(retryAfterS) && retryAfterS > 0
            ? retryAfterS * 1000
            : this.backoff(attempt);
        // Leave room for the retry itself; a wait that cannot fit is
        // surfaced as the 429 it is (the caller reports "nothing was sent").
        if (now() + delay + 5_000 < deadline) {
          await this.sleepFn(delay);
          continue;
        }
      }
      throw new GraphApiError(r.status, url, text, full);
    }
  }

  private backoff(attempt: number): number {
    return Math.min(60_000, 1000 * 2 ** attempt) + this.random() * 250;
  }
}
