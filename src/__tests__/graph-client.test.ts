/**
 * Ported from legacy `src/__tests__/ms365-client.test.ts` graphFetch suite
 * (retry/backoff behavior), reshaped onto GraphClient's host-fetch idiom
 * (see google-docs-kia-connector's client.test.ts for the pattern this
 * mirrors).
 */
import { GraphApiError, GraphClient, Ms365AuthError } from '../graph-client';
import { instantClock } from '../testing/harness';

function hostFetchSeq(
  responses: Array<{ status: number; body: unknown; headers?: Record<string, string> }>,
) {
  let i = 0;
  const calls: number[] = [];
  const fetchFn = async () => {
    calls.push(Date.now());
    const r = responses[Math.min(i, responses.length - 1)];
    i++;
    return {
      status: r.status,
      statusText: '',
      headers: r.headers ?? {},
      body: new TextEncoder().encode(JSON.stringify(r.body)),
    };
  };
  return { fetchFn, callCount: () => i };
}

describe('GraphClient.request', () => {
  it('returns JSON on 200', async () => {
    const { fetchFn } = hostFetchSeq([{ status: 200, body: { id: 'oid' } }]);
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).resolves.toEqual({
      id: 'oid',
    });
  });

  it('retries 429 honoring Retry-After (seconds), tracked via the injectable sleep', async () => {
    const sleeps: number[] = [];
    const { fetchFn } = hostFetchSeq([
      { status: 429, body: '', headers: { 'retry-after': '1' } },
      { status: 200, body: { id: 'oid' } },
    ]);
    const client = new GraphClient({
      fetch: fetchFn,
      getToken: async () => 'tok',
      sleep: async (ms) => {
        sleeps.push(ms);
      },
      random: () => 0,
    });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).resolves.toEqual({
      id: 'oid',
    });
    expect(sleeps).toEqual([1000]);
  });

  it('throws after exhausting retries on persistent 500', async () => {
    const { fetchFn, callCount } = hostFetchSeq([{ status: 500, body: 'oops' }]);
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).rejects.toThrow(/500/);
    expect(callCount()).toBe(5); // initial + 4 retries
  });

  it('does not retry 4xx other than 429', async () => {
    const { fetchFn, callCount } = hostFetchSeq([{ status: 404, body: 'nope' }]);
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).rejects.toThrow(
      GraphApiError,
    );
    expect(callCount()).toBe(1);
  });

  it('never retries a 401 and throws Ms365AuthError', async () => {
    const { fetchFn, callCount } = hostFetchSeq([{ status: 401, body: 'nope' }]);
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).rejects.toThrow(
      Ms365AuthError,
    );
    expect(callCount()).toBe(1);
  });

  it('retries a network error up to MAX_RETRIES then throws it', async () => {
    let calls = 0;
    const fetchFn = async () => {
      calls++;
      throw new Error('ECONNRESET');
    };
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 'tok', ...instantClock });
    await expect(client.request(`https://graph.microsoft.com/v1.0/me`)).rejects.toThrow(
      'ECONNRESET',
    );
    expect(calls).toBe(5);
  });
});
