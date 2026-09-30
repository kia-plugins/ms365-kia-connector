import type { DocumentInput, SendIntent } from '@kiagent/connector-sdk';
import { GRAPH_BASE } from '../graph-api';
import { GraphApiError, GraphClient } from '../graph-client';
import type { GraphMessage } from '../parser';
import { replyTargets } from '../reply-target';
import { createMs365Sender } from '../sender';
import { createMs365Source } from '../source';
import { toDocument } from '../to-document';
import { graphMsg, instantClock, makeHost } from '../testing/harness';

const ME = 'me@contoso.com';
const who = (address: string, name = '') => ({ emailAddress: { address, name } });

function msg(id: string, over: Partial<GraphMessage>): GraphMessage {
  return graphMsg({ id, conversationId: 'CONV', ...over });
}

describe('reply targets (mirrors resolve-gmail)', () => {
  it('reply goes to the last message not from you; reply-all to its From + To and Cc, never you', () => {
    const out = replyTargets(
      [
        msg('M1', { from: who('alice@contoso.com'), toRecipients: [who(ME)] }),
        msg('M2', {
          from: who('bob@contoso.com', 'Bob'),
          toRecipients: [who(ME), who('alice@contoso.com')],
          ccRecipients: [who('carol@contoso.com'), who('ME@contoso.com')],
        }),
      ],
      ME,
    );
    expect(out).toEqual({
      ref: { messageId: 'M2' },
      display: 'bob@contoso.com',
      to: ['bob@contoso.com'],
      cc: [],
      replyAll: {
        ref: { messageId: 'M2' },
        display: 'bob@contoso.com, alice@contoso.com; cc carol@contoso.com',
        to: ['bob@contoso.com', 'alice@contoso.com'],
        cc: ['carol@contoso.com'],
      },
    });
  });

  it("Reply-To wins over From, but Graph's empty Reply-To list means From", () => {
    const withReplyTo = replyTargets(
      [msg('M1', { from: who('noreply@list.com'), replyTo: [who('list@list.com')], toRecipients: [who(ME)] })],
      ME,
    );
    expect(withReplyTo?.to).toEqual(['list@list.com']);
    expect(withReplyTo?.replyAll?.to).toEqual(['list@list.com']);

    const emptyReplyTo = replyTargets(
      [msg('M1', { from: who('alice@contoso.com'), replyTo: [], toRecipients: [who(ME), who('bob@contoso.com')] })],
      ME,
    );
    expect(emptyReplyTo?.to).toEqual(['alice@contoso.com']);
    expect(emptyReplyTo?.replyAll?.to).toEqual(['alice@contoso.com', 'bob@contoso.com']);
  });

  it('a Reply-To pointing back at you falls back to From', () => {
    const out = replyTargets(
      [msg('M1', { from: who('alice@contoso.com'), replyTo: [who(ME)], toRecipients: [who(ME)] })],
      ME,
    );
    expect(out?.to).toEqual(['alice@contoso.com']);
  });

  it('an unsent draft never picks or adds a recipient', () => {
    const out = replyTargets(
      [
        msg('M1', { from: who('alice@contoso.com'), toRecipients: [who(ME)] }),
        msg('D1', { from: who(ME), toRecipients: [who('stranger@x.com')], isDraft: true }),
      ],
      ME,
    );
    expect(out?.ref).toEqual({ messageId: 'M1' });
    expect(out?.replyAll?.to).toEqual(['alice@contoso.com']);
    expect(JSON.stringify(out)).not.toContain('stranger');
  });

  it('when you wrote every message, reply goes to your last message’s recipients', () => {
    const out = replyTargets(
      [msg('M1', { from: who(ME), toRecipients: [who('dana@contoso.com')] })],
      ME,
    );
    expect(out?.to).toEqual(['dana@contoso.com']);
    expect(out?.ref).toEqual({ messageId: 'M1' });
  });

  it('no target at all when nobody but you is on the thread', () => {
    expect(replyTargets([msg('M1', { from: who(ME), toRecipients: [who(ME)] })], ME)).toBeUndefined();
  });

  it('toDocument stores the targets under metadata.outbound, keyed by the raw Graph id (not the RFC Message-ID)', () => {
    const out = toDocument({
      conversationId: 'CONV',
      tenantKind: 'work',
      scopeRootId: null,
      selfAddress: ME,
      messages: [msg('AAMk-immutable', { internetMessageId: '<rfc@x>', from: who('alice@contoso.com'), toRecipients: [who(ME)] })],
    }) as DocumentInput;
    expect((out.metadata.outbound as { ref: unknown }).ref).toEqual({ messageId: 'AAMk-immutable' });
  });
});

describe('ms365 sender', () => {
  const REPLY: SendIntent = {
    accountId: 'A' as never,
    kind: 'reply',
    outboundRef: { messageId: 'AAMk/1+x=' },
    to: ['bob@contoso.com'],
    cc: ['carol@contoso.com'],
    bodyMarkdown: 'Thanks,\nEd',
  };

  function sender(responses: Array<{ status: number; body?: unknown; headers?: Record<string, string> }>) {
    const calls: Array<{ url: string; init: { method?: string; body?: string; timeoutMs?: number } }> = [];
    let i = 0;
    const fetchFn = async (url: string, init: never) => {
      calls.push({ url, init });
      const r = responses[Math.min(i++, responses.length - 1)];
      return {
        status: r.status,
        statusText: '',
        headers: r.headers ?? {},
        body: new TextEncoder().encode(r.body === undefined ? '' : JSON.stringify(r.body)),
      };
    };
    return { s: createMs365Sender(makeHost(fetchFn as never) as never, instantClock), calls };
  }
  const ctx = { credentials: { accessToken: 'T' } };

  it('replies on the stored message with exactly the confirmed recipients, as plain text', async () => {
    const { s, calls } = sender([{ status: 202 }]);
    await expect(s.send(REPLY, ctx)).resolves.toEqual({});
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe(`${GRAPH_BASE}/me/messages/AAMk%2F1%2Bx%3D/reply`);
    expect(calls[0].init.method).toBe('POST');
    expect(JSON.parse(calls[0].init.body!)).toEqual({
      message: {
        toRecipients: [{ emailAddress: { address: 'bob@contoso.com' } }],
        ccRecipients: [{ emailAddress: { address: 'carol@contoso.com' } }],
        body: { contentType: 'Text', content: 'Thanks,\nEd' },
      },
    });
    // Bounded by the send deadline (kiagent-core gives up at 60 s).
    expect(calls[0].init.timeoutMs).toBeLessThanOrEqual(40_000);
  });

  it('new mail goes through sendMail and is saved to Sent Items', async () => {
    const { s, calls } = sender([{ status: 202 }]);
    await s.send(
      { accountId: 'A' as never, kind: 'new', to: ['dana@contoso.com'], subject: 'Hi', bodyMarkdown: 'Hello' },
      ctx,
    );
    expect(calls[0].url).toBe(`${GRAPH_BASE}/me/sendMail`);
    expect(JSON.parse(calls[0].init.body!)).toEqual({
      message: {
        subject: 'Hi',
        body: { contentType: 'Text', content: 'Hello' },
        toRecipients: [{ emailAddress: { address: 'dana@contoso.com' } }],
        ccRecipients: [],
      },
      saveToSentItems: true,
    });
  });

  it.each([
    [500, {}],
    [503, {}],
  ])('never retries a %s — Graph may already have accepted the message', async (status, body) => {
    const { s, calls } = sender([{ status, body }, { status: 202 }]);
    const err = await s.send(REPLY, ctx).then(() => null, (e: Error) => e);
    expect(calls).toHaveLength(1);
    // Not one of the "nothing was sent" markers.
    expect(err?.message).toMatch(/^graph 5\d\d /);
  });

  it('never retries a network error', async () => {
    let n = 0;
    const s = createMs365Sender(
      makeHost((async () => {
        n += 1;
        throw new Error('socket hang up');
      }) as never) as never,
      instantClock,
    );
    await expect(s.send(REPLY, ctx)).rejects.toThrow('socket hang up');
    expect(n).toBe(1);
  });

  it('retries a 429 that fits the deadline; one that cannot fit is "rate-limited", nothing sent', async () => {
    const ok = sender([{ status: 429, headers: { 'retry-after': '2' } }, { status: 202 }]);
    await expect(ok.s.send(REPLY, ctx)).resolves.toEqual({});
    expect(ok.calls).toHaveLength(2);

    const long = sender([{ status: 429, headers: { 'retry-after': '120' } }, { status: 202 }]);
    await expect(long.s.send(REPLY, ctx)).rejects.toThrow(/^rate-limited: /);
    expect(long.calls).toHaveLength(1);
  });

  it.each([
    [401, { error: { code: 'InvalidAuthenticationToken' } }, /reconnect the account in Settings$/],
    [403, { error: { code: 'ErrorAccessDenied' } }, /reconnect it in Settings to grant send permission$/],
    [404, { error: { code: 'ErrorItemNotFound' } }, /^not sent: the original message no longer exists/],
    [400, { error: { code: 'ErrorInvalidRecipients' } }, /^not sent: Microsoft 365 rejected the message \(ErrorInvalidRecipients\)$/],
  ])('maps %s to the outbound wording', async (status, body, rx) => {
    const { s } = sender([{ status, body }]);
    await expect(s.send(REPLY, ctx)).rejects.toThrow(rx);
  });

  it('a 403 that is not an access denial stays unclassified', async () => {
    const { s } = sender([{ status: 403, body: { error: { code: 'ErrorQuotaExceeded' } } }]);
    await expect(s.send(REPLY, ctx)).rejects.toBeInstanceOf(GraphApiError);
  });

  it('refuses without credentials or a target, before any request', async () => {
    const { s, calls } = sender([{ status: 202 }]);
    await expect(s.send(REPLY, { credentials: null })).rejects.toThrow(/reconnect the account in Settings/);
    await expect(s.send({ ...REPLY, outboundRef: {} }, ctx)).rejects.toThrow(/^not sent: /);
    await expect(s.send({ ...REPLY, to: [] }, ctx)).rejects.toThrow(/^not sent: /);
    expect(calls).toHaveLength(0);
  });
});

describe('GraphClient.post', () => {
  it('bounds the whole send by one deadline: waits add up, and the one that would cross it is not slept', async () => {
    let t = 0;
    const slept: number[] = [];
    const client = new GraphClient({
      fetch: (async () => ({
        status: 429,
        statusText: '',
        headers: { 'retry-after': '20' },
        body: new Uint8Array(),
      })) as never,
      getToken: async () => 't',
      sleep: async (ms) => {
        slept.push(ms);
        t += ms;
      },
    });
    await expect(
      client.post(`${GRAPH_BASE}/me/sendMail`, {}, { deadlineMs: 40_000, now: () => t }),
    ).rejects.toMatchObject({ status: 429 });
    // 20 s fits (20 + 5 < 40); a second 20 s from t=20 would not.
    expect(slept).toEqual([20_000]);
  });
});

describe('source descriptor', () => {
  it("declares compose: 'email' so new mail can be drafted from ms365 accounts", () => {
    const src = createMs365Source(makeHost((async () => ({})) as never));
    expect((src.descriptor as { compose?: string }).compose).toBe('email');
  });
});
