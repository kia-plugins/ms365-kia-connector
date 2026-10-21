import type { SendIntent } from '@kiagent/connector-sdk';
import {
  DRAFT_RECONNECT,
  SIMPLE_ATTACHMENT_MAX_BYTES,
  UPLOAD_CHUNK_BYTES,
  createMs365Drafts,
} from '../drafts';
import { GRAPH_BASE } from '../graph-api';
import { createMs365Sender } from '../sender';
import { instantClock, makeHost } from '../testing/harness';

type Res = { status: number; body?: unknown };
type Call = { method: string; url: string; headers: Record<string, string>; body?: string | Uint8Array };

/** A fake Graph: the first route whose method + url pattern match answers. */
function graph(routes: Array<[string, RegExp, Res | ((c: Call) => Res)]>) {
  const calls: Call[] = [];
  const fetchFn = async (
    url: string,
    init: { method?: string; headers?: Record<string, string>; body?: string | Uint8Array },
  ) => {
    const c: Call = { method: init.method ?? 'GET', url, headers: init.headers ?? {}, body: init.body };
    calls.push(c);
    const hit = routes.find(([m, rx]) => m === c.method && rx.test(url.replace(GRAPH_BASE, '')));
    const r = hit ? (typeof hit[2] === 'function' ? hit[2](c) : hit[2]) : { status: 599 };
    return {
      status: r.status,
      statusText: '',
      headers: {},
      body: new TextEncoder().encode(r.body === undefined ? '' : JSON.stringify(r.body)),
    };
  };
  const d = createMs365Drafts(makeHost(fetchFn as never) as never, instantClock);
  const json = (c: Call) => JSON.parse(c.body as string);
  return { d, calls, json };
}

const ctx = { credentials: { accessToken: 'T' } };
const REPLY: SendIntent = {
  accountId: 'A' as never,
  kind: 'reply',
  outboundRef: { messageId: 'M1' },
  to: ['bob@contoso.com'],
  cc: ['carol@contoso.com'],
  subject: 'Re: Plan',
  bodyMarkdown: 'Thanks,\nEd',
};
const NEW: SendIntent = {
  accountId: 'A' as never,
  kind: 'new',
  to: ['bob@contoso.com'],
  subject: 'Plan',
  bodyMarkdown: 'Hi',
};
const REF = { accountId: 'A' as never, id: 'D1' };
const file = (filename: string, n: number) => ({
  filename,
  mime: 'application/pdf',
  contentBase64: Buffer.alloc(n, 1).toString('base64'),
});

describe('ms365 drafts: create', () => {
  it('a reply is createReply, then the confirmed recipients and Kia text', async () => {
    const { d, calls, json } = graph([
      ['POST', /^\/me\/messages\/M1\/createReply$/, { status: 201, body: { id: 'D1', conversationId: 'C1' } }],
      ['PATCH', /^\/me\/messages\/D1$/, { status: 200, body: {} }],
    ]);
    await expect(d.create({ ...REPLY, bcc: ['x@contoso.com'] }, ctx)).resolves.toEqual({
      accountId: 'A',
      id: 'D1',
      threadId: 'C1',
    });
    expect(calls.map((c) => c.method)).toEqual(['POST', 'PATCH']);
    expect(json(calls[1])).toEqual({
      toRecipients: [{ emailAddress: { address: 'bob@contoso.com' } }],
      ccRecipients: [{ emailAddress: { address: 'carol@contoso.com' } }],
      bccRecipients: [{ emailAddress: { address: 'x@contoso.com' } }],
      subject: 'Re: Plan',
      body: { contentType: 'Text', content: 'Thanks,\nEd' },
    });
    expect(calls[0].headers.prefer).toContain('ImmutableId');
  });

  it('new mail is POST /me/messages', async () => {
    const { d, calls, json } = graph([['POST', /^\/me\/messages$/, { status: 201, body: { id: 'D2' } }]]);
    await expect(d.create(NEW, ctx)).resolves.toEqual({ accountId: 'A', id: 'D2' });
    expect(json(calls[0])).toMatchObject({
      subject: 'Plan',
      body: { contentType: 'Text', content: 'Hi' },
      toRecipients: [{ emailAddress: { address: 'bob@contoso.com' } }],
    });
  });

  it('a formatted draft is written as HTML (platform 2.15.0)', async () => {
    const { d, calls, json } = graph([['POST', /^\/me\/messages$/, { status: 201, body: { id: 'D3' } }]]);
    await d.create({ ...NEW, bodyHtml: '<div><p><b>Hi</b></p></div>' }, ctx);
    expect(json(calls[0]).body).toEqual({ contentType: 'HTML', content: '<div><p><b>Hi</b></p></div>' });
  });

  it('a small file is one POST; a large one goes through an upload session in 320 KiB-multiple chunks, without the token', async () => {
    const big = UPLOAD_CHUNK_BYTES + 10;
    expect(big).toBeGreaterThan(SIMPLE_ATTACHMENT_MAX_BYTES);
    const { d, calls, json } = graph([
      ['POST', /^\/me\/messages$/, { status: 201, body: { id: 'D2' } }],
      ['POST', /^\/me\/messages\/D2\/attachments$/, { status: 201, body: {} }],
      [
        'POST',
        /^\/me\/messages\/D2\/attachments\/createUploadSession$/,
        { status: 201, body: { uploadUrl: 'https://upload.example/s1' } },
      ],
      ['PUT', /^https:\/\/upload\.example\/s1$/, { status: 200, body: {} }],
    ]);
    await d.create({ ...NEW, attachments: [file('a.pdf', 10), file('big.pdf', big)] }, ctx);
    expect(json(calls[1])).toEqual({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'a.pdf',
      contentType: 'application/pdf',
      contentBytes: Buffer.alloc(10, 1).toString('base64'),
    });
    expect(json(calls[2])).toEqual({
      AttachmentItem: { attachmentType: 'file', name: 'big.pdf', size: big, contentType: 'application/pdf' },
    });
    const puts = calls.filter((c) => c.method === 'PUT');
    expect(puts.map((c) => c.headers['content-range'])).toEqual([
      `bytes 0-${UPLOAD_CHUNK_BYTES - 1}/${big}`,
      `bytes ${UPLOAD_CHUNK_BYTES}-${big - 1}/${big}`,
    ]);
    expect(UPLOAD_CHUNK_BYTES % (320 * 1024)).toBe(0);
    expect(puts.every((c) => !c.headers.authorization)).toBe(true);
    expect((puts[0].body as Uint8Array).byteLength).toBe(UPLOAD_CHUNK_BYTES);
  });

  it('a file that fails deletes the half-made draft and says why', async () => {
    const { d, calls } = graph([
      ['POST', /^\/me\/messages$/, { status: 201, body: { id: 'D2' } }],
      ['POST', /attachments$/, { status: 400, body: { error: { code: 'ErrorAttachmentSizeLimitExceeded' } } }],
      ['DELETE', /^\/me\/messages\/D2$/, { status: 204 }],
    ]);
    await expect(d.create({ ...NEW, attachments: [file('a.pdf', 10)] }, ctx)).rejects.toThrow(
      /ErrorAttachmentSizeLimitExceeded/,
    );
    expect(calls.at(-1)).toMatchObject({ method: 'DELETE' });
  });

  it('an account connected with Mail.Read only (403 on create) asks for a reconnect', async () => {
    const { d } = graph([
      ['POST', /^\/me\/messages$/, { status: 403, body: { error: { code: 'ErrorAccessDenied' } } }],
    ]);
    await expect(d.create(NEW, ctx)).rejects.toThrow(DRAFT_RECONNECT);
  });

  it('refuses without credentials, before any request', async () => {
    const { d, calls } = graph([]);
    await expect(d.create(NEW, { credentials: null })).rejects.toThrow(/reconnect the account in Settings/);
    expect(calls).toHaveLength(0);
  });
});

describe('ms365 drafts: update', () => {
  it('patches only the given fields and swaps files by name', async () => {
    const { d, calls, json } = graph([
      ['PATCH', /^\/me\/messages\/D1$/, { status: 200, body: {} }],
      [
        'GET',
        /^\/me\/messages\/D1\/attachments\?/,
        {
          status: 200,
          body: {
            value: [
              { id: 'att-1', name: 'old.pdf' },
              { id: 'att-2', name: 'logo.png', isInline: true },
              { id: 'att-3', name: 'keep.pdf' },
            ],
          },
        },
      ],
      ['DELETE', /^\/me\/messages\/D1\/attachments\/att-1$/, { status: 204 }],
      ['POST', /^\/me\/messages\/D1\/attachments$/, { status: 201, body: {} }],
    ]);
    await expect(
      d.update(REF, { bodyMarkdown: 'New text', removeAttachments: ['old.pdf'], attachments: [file('new.pdf', 5)] }, ctx),
    ).resolves.toEqual(REF);
    expect(json(calls[0])).toEqual({ body: { contentType: 'Text', content: 'New text' } });
    expect(calls.filter((c) => c.method === 'DELETE').map((c) => c.url)).toEqual([
      `${GRAPH_BASE}/me/messages/D1/attachments/att-1`,
    ]);
    expect(json(calls.at(-1)!).name).toBe('new.pdf');
  });

  it('a body patch with HTML writes HTML', async () => {
    const { d, calls, json } = graph([['PATCH', /^\/me\/messages\/D1$/, { status: 200, body: {} }]]);
    await d.update(REF, { bodyMarkdown: '**New**', bodyHtml: '<p><strong>New</strong></p>' }, ctx);
    expect(json(calls[0])).toEqual({ body: { contentType: 'HTML', content: '<p><strong>New</strong></p>' } });
  });

  it('a files-only change sends no PATCH', async () => {
    const { d, calls } = graph([['POST', /attachments$/, { status: 201, body: {} }]]);
    await d.update(REF, { attachments: [file('n.pdf', 1)] }, ctx);
    expect(calls.map((c) => c.method)).toEqual(['POST']);
  });

  it('retries a 5xx on the idempotent PATCH', async () => {
    let n = 0;
    const { d, calls } = graph([
      ['PATCH', /^\/me\/messages\/D1$/, () => (n++ === 0 ? { status: 503 } : { status: 200, body: {} })],
    ]);
    await d.update(REF, { subject: 'X' }, ctx);
    expect(calls).toHaveLength(2);
  });
});

describe('ms365 drafts: read', () => {
  const folders: [string, RegExp, Res] = [
    'GET',
    /^\/me\/mailFolders\/deleteditems/,
    { status: 200, body: { id: 'DELETED', displayName: 'Deleted Items' } },
  ];

  it('a draft reads as text with its files (inline images left out)', async () => {
    const { d, calls } = graph([
      [
        'GET',
        /^\/me\/messages\/D1\?/,
        {
          status: 200,
          body: {
            isDraft: true,
            parentFolderId: 'DRAFTS',
            conversationId: 'C1',
            subject: 'Re: Plan',
            toRecipients: [{ emailAddress: { address: 'bob@contoso.com' } }],
            ccRecipients: [],
            bccRecipients: [{ emailAddress: { address: 'x@contoso.com' } }],
            body: { contentType: 'text', content: 'Thanks,\r\nEd' },
          },
        },
      ],
      folders,
      [
        'GET',
        /^\/me\/messages\/D1\/attachments\?/,
        {
          status: 200,
          body: {
            value: [
              { name: 'a.pdf', contentType: 'application/pdf', size: 10 },
              { name: 'logo.png', contentType: 'image/png', size: 900, isInline: true },
            ],
          },
        },
      ],
    ]);
    await expect(d.read(REF, ctx)).resolves.toEqual({
      state: 'draft',
      ref: { ...REF, threadId: 'C1' },
      content: {
        to: ['bob@contoso.com'],
        cc: [],
        bcc: ['x@contoso.com'],
        subject: 'Re: Plan',
        bodyText: 'Thanks,\nEd',
        attachments: [{ filename: 'a.pdf', mime: 'application/pdf', size: 10 }],
      },
    });
    // The body as Outlook holds it — no text conversion asked for.
    expect(calls[0].headers.prefer).not.toContain('body-content-type');
    expect(calls[0].headers.prefer).toContain('ImmutableId');
  });

  it('an HTML draft reads as its HTML and that HTML as text', async () => {
    const { d } = graph([
      [
        'GET',
        /^\/me\/messages\/D1\?/,
        {
          status: 200,
          body: {
            isDraft: true,
            parentFolderId: 'DRAFTS',
            subject: 'Plan',
            toRecipients: [],
            body: { contentType: 'html', content: '<html><body><h2>Plan</h2>\r\n<p><b>one</b></p></body></html>' },
          },
        },
      ],
      folders,
      ['GET', /^\/me\/messages\/D1\/attachments\?/, { status: 200, body: { value: [] } }],
    ]);
    const r = await d.read(REF, ctx);
    expect(r.state).toBe('draft');
    const c = (r as { content: { bodyText: string; bodyHtml?: string } }).content;
    expect(c.bodyHtml).toBe('<html><body><h2>Plan</h2>\n<p><b>one</b></p></body></html>');
    expect(c.bodyText).toBe('## Plan\n\n**one**');
  });

  it('isDraft false → sent; in Deleted Items → deleted; 404 → gone', async () => {
    const sent = graph([
      ['GET', /^\/me\/messages\/D1\?/, { status: 200, body: { isDraft: false, internetMessageId: '<m@x>' } }],
    ]);
    await expect(sent.d.read(REF, ctx)).resolves.toEqual({ state: 'sent', externalMessageId: '<m@x>' });

    const deleted = graph([
      ['GET', /^\/me\/messages\/D1\?/, { status: 200, body: { isDraft: true, parentFolderId: 'DELETED' } }],
      folders,
    ]);
    await expect(deleted.d.read(REF, ctx)).resolves.toEqual({ state: 'deleted' });

    const gone = graph([
      ['GET', /^\/me\/messages\/D1\?/, { status: 404, body: { error: { code: 'ErrorItemNotFound' } } }],
    ]);
    await expect(gone.d.read(REF, ctx)).resolves.toEqual({ state: 'gone' });
  });

  it('a throttled Deleted Items lookup fails this read only; the next one asks again', async () => {
    let throttled = true;
    const { d, calls } = graph([
      ['GET', /^\/me\/messages\/D1\?/, { status: 200, body: { isDraft: true, parentFolderId: 'DELETED' } }],
      ['GET', /^\/me\/mailFolders\/deleteditems/, () =>
        throttled ? { status: 429 } : { status: 200, body: { id: 'DELETED' } }],
    ]);
    await expect(d.read(REF, ctx)).rejects.toThrow(/graph 429/);
    throttled = false;
    await expect(d.read(REF, ctx)).resolves.toEqual({ state: 'deleted' });
    await d.read(REF, ctx);
    // five throttled tries, one found, then remembered
    expect(calls.filter((c) => /mailFolders/.test(c.url))).toHaveLength(6);
  });

  it('a mailbox-level 404 is not "gone"', async () => {
    const { d } = graph([
      ['GET', /^\/me\/messages\/D1\?/, { status: 404, body: { error: { code: 'ResourceNotFound' } } }],
    ]);
    await expect(d.read(REF, ctx)).rejects.toThrow(/graph 404/);
  });
});

describe('ms365 drafts: send and remove', () => {
  it('send is POST …/send, never retried after a 5xx', async () => {
    const { d, calls } = graph([['POST', /^\/me\/messages\/D1\/send$/, { status: 502 }]]);
    await expect(d.send(REF, ctx)).rejects.toThrow(/graph 502/);
    expect(calls).toHaveLength(1);

    const ok = graph([['POST', /^\/me\/messages\/D1\/send$/, { status: 202 }]]);
    await expect(ok.d.send(REF, ctx)).resolves.toEqual({});
  });

  it('a send of a draft that is gone is provably not sent', async () => {
    const { d } = graph([
      ['POST', /\/send$/, { status: 404, body: { error: { code: 'ErrorItemNotFound' } } }],
    ]);
    await expect(d.send(REF, ctx)).rejects.toThrow(/^not sent: the Outlook draft no longer exists/);
  });

  it('remove is DELETE, and an already-gone draft is fine', async () => {
    const { d, calls } = graph([
      ['DELETE', /^\/me\/messages\/D1$/, { status: 404, body: { error: { code: 'ErrorItemNotFound' } } }],
    ]);
    await expect(d.remove(REF, ctx)).resolves.toBeUndefined();
    expect(calls).toHaveLength(1);
  });
});

it('the sender carries the draft mirror', () => {
  expect(createMs365Sender(makeHost((async () => ({})) as never) as never).drafts).toBeDefined();
});
