/**
 * listAddressedTo (Email Kia): the account's own requests come from Sent
 * Items with their body; everything else addressed to you+kia@ anywhere in
 * the mailbox (Junk and Deleted included) is listed as metadata only, so the
 * assistant can log it — its body never leaves Graph.
 */
import { listAddressedTo } from '../addressed';
import { GraphClient } from '../graph-client';
import { createMs365Source } from '../source';
import {
  graphFetch,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
  type HostResponse,
} from '../testing/harness';

const SENT = 'SENTID';
const KIA = 'me+kia@company.com';
const to = (...a: string[]) => a.map((address) => ({ emailAddress: { address } }));
const from = (address: string) => ({ emailAddress: { address } });

type Row = Record<string, unknown>;
interface Fx {
  sent?: Row[][];
  all?: Row[][];
  full?: Record<string, Row>;
}

/** Routes the three Graph calls listAddressedTo makes; records full GETs. */
function world(fx: Fx) {
  const fetched: string[] = [];
  const queries: URLSearchParams[] = [];
  const paged = (url: URL, pages: Row[][] = [[]]): HostResponse => {
    const i = Number(url.searchParams.get('pageToken') ?? 0);
    const body: Row = { value: pages[i] ?? [] };
    if (i + 1 < pages.length) {
      const next = new URL(url.toString());
      next.searchParams.set('pageToken', String(i + 1));
      body['@odata.nextLink'] = next.toString();
    }
    return jsonRes(200, body);
  };
  const { fetchFn } = graphFetch({
    folders: { top: [], wellKnown: { sentitems: { id: SENT, displayName: 'Sent Items', childFolderCount: 0 } } },
    custom: (url) => {
      const p = url.pathname;
      if (p === '/v1.0/me/mailFolders/sentitems/messages') {
        queries.push(url.searchParams);
        return paged(url, fx.sent);
      }
      if (p === '/v1.0/me/messages' && !url.searchParams.get('$filter')?.includes('conversationId')) {
        queries.push(url.searchParams);
        return paged(url, fx.all);
      }
      const one = /^\/v1\.0\/me\/messages\/([^/]+)$/.exec(p);
      if (one) {
        const id = decodeURIComponent(one[1]);
        fetched.push(id);
        const m = fx.full?.[id];
        return m ? jsonRes(200, m) : jsonRes(404, { error: { code: 'ErrorItemNotFound' } });
      }
      return undefined;
    },
  });
  const client = new GraphClient({ fetch: fetchFn, getToken: async () => 't', ...instantClock });
  return { client, fetchFn, fetched, queries };
}

const request = (over: Row = {}): Row => ({
  id: 'S1',
  internetMessageId: '<q1@x>',
  toRecipients: to(KIA),
  sentDateTime: '2026-10-04T10:00:00Z',
  from: from('me@company.com'),
  sender: from('me@company.com'),
  subject: 'Prep',
  conversationId: 'C1',
  ...over,
});
const fullOf = (over: Row = {}): Row => ({
  internetMessageHeaders: [{ name: 'Message-ID', value: '<q1@x>' }],
  body: { contentType: 'text', content: 'Prepare me' },
  conversationId: 'C1',
  ...over,
});
const audit = (over: Row): Row => ({
  isDraft: false,
  toRecipients: to(KIA),
  from: from('me@company.com'),
  subject: 'x',
  receivedDateTime: '2026-10-04T10:00:01Z',
  ...over,
});

describe('listAddressedTo', () => {
  it('a sent request comes with its body; a self-copy and drafts are skipped; foreign mail is metadata only', async () => {
    const { client, fetched } = world({
      sent: [[request()]],
      all: [
        [
          audit({ id: 'I1', internetMessageId: '<q1@x>', parentFolderId: 'INBOX' }),
          audit({
            id: 'J1',
            internetMessageId: '<e@x>',
            parentFolderId: 'JUNK',
            from: from('Evil@Y.com'),
            subject: 'hi',
            receivedDateTime: '2026-10-04T10:00:02Z',
          }),
          audit({ id: 'D1', internetMessageId: '<d@x>', parentFolderId: 'DRAFTS', isDraft: true }),
          audit({ id: 'S1b', internetMessageId: '<q2@x>', parentFolderId: SENT }),
          audit({ id: 'O1', internetMessageId: '<o@x>', parentFolderId: 'INBOX', toRecipients: to('me@company.com') }),
        ],
      ],
      full: { S1: fullOf() },
    });
    const rows = await listAddressedTo(client, { toAddress: KIA, since: 0 }, SENT);
    expect(rows).toEqual([
      {
        providerId: 'S1',
        messageId: '<q1@x>',
        references: [],
        providerThreadId: 'C1',
        sent: true,
        from: 'me@company.com',
        subject: 'Prep',
        date: Date.parse('2026-10-04T10:00:00Z'),
        to: [KIA],
        cc: [],
        text: 'Prepare me',
        ownChannel: false,
        automated: false,
      },
      {
        providerId: 'J1',
        messageId: '<e@x>',
        references: [],
        sent: false,
        from: 'evil@y.com',
        subject: 'hi',
        date: Date.parse('2026-10-04T10:00:02Z'),
        to: [KIA],
        cc: [],
        text: '',
        ownChannel: false,
        automated: false,
      },
    ]);
    // Only the account's own request is read in full.
    expect(fetched).toEqual(['S1']);
  });

  it('threading, channel and automation come from the headers; addresses are lowercased', async () => {
    const { client } = world({
      sent: [
        [
          request({
            toRecipients: to('Me+Kia@Company.com', 'Bob@x.com'),
            ccRecipients: to('Carol@X.com'),
            from: from('Me@Company.com'),
            sender: from('Assistant@Company.com'),
          }),
        ],
      ],
      full: {
        S1: fullOf({
          internetMessageHeaders: [
            { name: 'In-Reply-To', value: '<a1@kia.local>' },
            { name: 'References', value: '<r0@x> <a1@kia.local>' },
            { name: 'X-Kia-Channel', value: 'owner' },
            { name: 'Auto-Submitted', value: 'auto-replied' },
          ],
        }),
      },
    });
    const [row] = await listAddressedTo(client, { toAddress: KIA, since: 0 }, SENT);
    expect(row).toMatchObject({
      inReplyTo: '<a1@kia.local>',
      references: ['<r0@x>', '<a1@kia.local>'],
      from: 'me@company.com',
      sender: 'assistant@company.com',
      to: [KIA, 'bob@x.com'],
      cc: ['carol@x.com'],
      ownChannel: true,
      automated: true,
    });
  });

  it('each automation header on its own marks the request automated', async () => {
    const cases: Array<[string, string]> = [
      ['Precedence', 'bulk'],
      ['List-Id', '<l.x>'],
      ['X-Auto-Response-Suppress', 'OOF, AutoReply'],
      ['X-Autoreply', 'yes'],
    ];
    for (const [name, value] of cases) {
      const { client } = world({
        sent: [[request()]],
        full: { S1: fullOf({ internetMessageHeaders: [{ name, value }] }) },
      });
      // eslint-disable-next-line no-await-in-loop
      const [row] = await listAddressedTo(client, { toAddress: KIA, since: 0 }, SENT);
      expect([name, row.automated]).toEqual([name, true]);
    }
  });

  it('sent mail not To the kia address is never read; pages are followed; the body is capped', async () => {
    const { client, fetched } = world({
      sent: [[request({ id: 'X1', toRecipients: to('bob@x.com') })], [request()]],
      full: { S1: fullOf({ body: { contentType: 'text', content: 'a'.repeat(40_000) } }) },
    });
    const rows = await listAddressedTo(client, { toAddress: KIA, since: 0 }, SENT);
    expect(rows.map((r) => r.providerId)).toEqual(['S1']);
    expect(rows[0].text).toHaveLength(32 * 1024);
    expect(fetched).toEqual(['S1']);
  });

  it('the sent leg starts at since, the audit leg at auditSince', async () => {
    const { client, queries } = world({});
    await listAddressedTo(
      client,
      { toAddress: KIA, since: Date.parse('2026-10-03T00:00:00Z'), auditSince: Date.parse('2026-10-04T09:00:00Z') },
      SENT,
    );
    expect(queries.map((q) => q.get('$filter'))).toEqual([
      'sentDateTime ge 2026-10-03T00:00:00.000Z',
      'receivedDateTime ge 2026-10-04T09:00:00.000Z',
    ]);
  });
});

describe('source.listAddressedTo', () => {
  it('resolves Sent Items and lists through it', async () => {
    const { fetchFn } = world({ sent: [[request()]], full: { S1: fullOf() } });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    const rows = await source.listAddressedTo!(makeSession().session, { toAddress: KIA, since: 0 });
    expect(rows.map((r) => [r.providerId, r.sent])).toEqual([['S1', true]]);
  });

  it('throws when the mailbox has no Sent Items folder', async () => {
    const { fetchFn } = graphFetch({ folders: { top: [], wellKnown: {} } });
    const source = createMs365Source(makeHost(fetchFn), instantClock);
    await expect(source.listAddressedTo!(makeSession().session, { toAddress: KIA, since: 0 })).rejects.toThrow(
      /Sent Items folder not found/,
    );
  });
});
