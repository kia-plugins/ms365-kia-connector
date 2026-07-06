/**
 * connect(auth) suite: platform-owned Microsoft OAuth (auth.oauth with
 * Graph-resource-only scopes), the /me identity fetch (mail ??
 * userPrincipalName), and the /organization tenant-kind probe that replaces
 * legacy's id_token `tid`-claim decode. Adapted from legacy
 * `src/__tests__/ms365-oauth.test.ts`'s identity/tenant-kind assertions,
 * which this port re-derives via live Graph calls instead of an id_token.
 */
import { createMs365Source, SCOPES } from '../source';
import { graphFetch, instantClock, makeAuth, makeHost } from '../testing/harness';

describe('connect', () => {
  it('oauth happy path: Graph-only scopes, statuses, identifier = mail, tenantKind = work', async () => {
    const { fetchFn, calls } = graphFetch({
      about: { mail: 'ed@corp.com', userPrincipalName: 'ed_corp.com#EXT#@tenant.onmicrosoft.com' },
      organization: { value: [{ id: 'org1' }] },
    });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth, statuses, getScopes } = makeAuth();

    const res = await source.connect(auth);

    expect(getScopes()).toEqual(SCOPES);
    expect(getScopes()).toEqual(['Mail.Read', 'User.Read']);
    expect(statuses).toEqual([
      'Waiting for Microsoft sign-in…',
      'Fetching Microsoft 365 profile…',
      'Checking Microsoft 365 account type…',
    ]);
    expect(res).toEqual({ identifier: 'ed@corp.com', config: { tenantKind: 'work' } });
    expect(calls.some((u) => u.includes('/v1.0/me?'))).toBe(true);
    expect(calls.some((u) => u.includes('/v1.0/organization'))).toBe(true);
  });

  it('falls back to userPrincipalName when mail is null', async () => {
    const { fetchFn } = graphFetch({
      about: { mail: null, userPrincipalName: 'alice@tenant.onmicrosoft.com' },
      organization: { value: [] },
    });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();
    const res = await source.connect(auth);
    expect(res.identifier).toBe('alice@tenant.onmicrosoft.com');
    expect(res.config).toEqual({ tenantKind: 'personal' });
  });

  it.each([401, 403, 404])(
    'treats a %d on /organization as a personal account, not an error',
    async (status) => {
      const { fetchFn } = graphFetch({
        about: { mail: 'me@outlook.com' },
        organization: { status, body: {} },
      });
      const source = createMs365Source(makeHost(fetchFn), instantClock);
      const { auth } = makeAuth();
      const res = await source.connect(auth);
      expect(res.config).toEqual({ tenantKind: 'personal' });
    },
  );

  it('classifies a non-empty /organization value as work', async () => {
    const { fetchFn } = graphFetch({
      about: { mail: 'me@corp.com' },
      organization: { value: [{ id: 'org-a' }, { id: 'org-b' }] },
    });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();
    const res = await source.connect(auth);
    expect(res.config).toEqual({ tenantKind: 'work' });
  });

  it('propagates an unexpected /organization failure (e.g. 500) rather than guessing personal', async () => {
    const { fetchFn } = graphFetch({
      about: { mail: 'me@corp.com' },
      organization: { status: 500, body: 'oops' },
    });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();
    await expect(source.connect(auth)).rejects.toThrow(/500/);
  });

  it('throws when oauth returns no accessToken, before any fetch', async () => {
    const { fetchFn, calls } = graphFetch({});
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth({ creds: { refreshToken: 'ms-test-refresh-deadbeef' } });
    await expect(source.connect(auth)).rejects.toThrow(/no access token/);
    expect(calls).toHaveLength(0);
  });

  it('throws when /me is missing both mail and userPrincipalName', async () => {
    const { fetchFn } = graphFetch({ about: { mail: null, userPrincipalName: undefined } });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const { auth } = makeAuth();
    await expect(source.connect(auth)).rejects.toThrow(/missing both mail and userPrincipalName/);
  });
});
