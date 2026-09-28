import type { Document, DocumentInput } from '@kiagent/connector-sdk';
import { GraphClient } from '../graph-client';
import { fetchConversationMessages, GRAPH_BASE } from '../graph-api';
import { createMs365Source } from '../source';
import { attachmentChildren, toDocument, type Ms365ThreadItem } from '../to-document';
import {
  graphFetch,
  graphMsg,
  instantClock,
  jsonRes,
  makeHost,
  makeSession,
  type HostResponse,
} from '../testing/harness';

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const FILE = '#microsoft.graph.fileAttachment';

const bytesRes = (body: Uint8Array): HostResponse =>
  ({ status: 200, headers: {}, body }) as HostResponse;

function thread(attachments: NonNullable<ReturnType<typeof graphMsg>['attachments']>): Ms365ThreadItem {
  return {
    conversationId: 'CONV',
    tenantKind: 'personal',
    scopeRootId: 'INBOX-ID',
    messages: [graphMsg({ id: 'IMMUTABLE-M1', conversationId: 'CONV', attachments })],
  };
}

describe('graph client', () => {
  it('asks for immutable ids on every request, merged with any other Prefer', async () => {
    const seen: Array<Record<string, string>> = [];
    const client = new GraphClient({
      fetch: (async (_url: string, init: { headers: Record<string, string> }) => {
        seen.push(init.headers);
        return jsonRes(200, {});
      }) as never,
      getToken: async () => 't',
      ...instantClock,
    });
    await client.request(`${GRAPH_BASE}/me`);
    await client.request(`${GRAPH_BASE}/me/messages`, {
      extraHeaders: { prefer: 'outlook.body-content-type="text"' },
    });
    expect(seen[0].prefer).toBe('IdType="ImmutableId"');
    expect(seen[1].prefer).toBe('IdType="ImmutableId", outlook.body-content-type="text"');
  });

  it("returns raw bytes for responseType 'bytes' (not UTF-8 decoded)", async () => {
    const payload = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xff, 0x00]);
    const client = new GraphClient({
      fetch: (async () => bytesRes(payload)) as never,
      getToken: async () => 't',
      ...instantClock,
    });
    const got = await client.request<Uint8Array>(`${GRAPH_BASE}/x/$value`, { responseType: 'bytes' });
    expect([...got]).toEqual([...payload]);
  });
});

describe('attachment enumeration', () => {
  it('expands attachment metadata on the conversation fetch (not gated on hasAttachments)', async () => {
    const { fetchFn, calls } = graphFetch({ conversations: { CONV: [graphMsg({ conversationId: 'CONV' })] } });
    const client = new GraphClient({ fetch: fetchFn, getToken: async () => 't', ...instantClock });
    await fetchConversationMessages(client, 'CONV');
    const url = new URL(calls[0]);
    expect(url.searchParams.get('$expand')).toBe('attachments($select=id,name,contentType,size,isInline)');
    expect(url.searchParams.get('$select')).not.toContain('contentBytes');
  });
});

describe('attachment children', () => {
  it('a file attachment becomes a bytes-less child keyed by immutable message id + name + size', () => {
    const out = toDocument(
      thread([{ '@odata.type': FILE, id: 'ATT-ID', name: 'offer.docx', contentType: DOCX, size: 33_630 }]),
    ) as DocumentInput[];
    expect(out).toHaveLength(2);
    expect(out[1]).toEqual({
      externalId: 'IMMUTABLE-M1#offer.docx#33630',
      type: 'attachment',
      title: 'offer.docx',
      markdown: null,
      metadata: { mime: DOCX, filename: 'offer.docx', sizeBytes: 33_630, messageId: 'IMMUTABLE-M1' },
      createdAt: expect.any(String),
      parent: { externalId: 'CONV', type: 'email.thread' },
      scopeRootId: 'INBOX-ID',
    });
    // The mutable attachment id is never hashed into the child.
    expect(JSON.stringify(out[1])).not.toContain('ATT-ID');
  });

  it('keeps a real inline screenshot, drops tiny inline images and non-file attachments', () => {
    const kids = attachmentChildren(
      thread([
        { '@odata.type': FILE, name: 'shot.png', contentType: 'image/png', size: 90_000, isInline: true },
        { '@odata.type': FILE, name: 'logo.png', contentType: 'image/png', size: 900, isInline: true },
        { '@odata.type': '#microsoft.graph.itemAttachment', name: 'fwd', size: 5_000 },
        { '@odata.type': '#microsoft.graph.referenceAttachment', name: 'link', size: 10 },
      ]),
    );
    expect(kids.map((k) => k.title)).toEqual(['shot.png']);
  });

  it('a thread without attachments stays a single document', () => {
    expect(Array.isArray(toDocument(thread([])))).toBe(false);
  });
});

describe('fetchBytes', () => {
  const doc = (meta: Record<string, unknown>) =>
    ({ id: 'd', type: 'attachment', metadata: meta }) as unknown as Document;
  const base = `${GRAPH_BASE}/me/messages/IMMUTABLE-M1/attachments`;

  function source(custom: (url: URL) => HostResponse | undefined) {
    const { fetchFn, calls } = graphFetch({ custom });
    const { session } = makeSession({ config: { tenantKind: 'personal' } });
    return { src: createMs365Source(makeHost(fetchFn), instantClock), session, calls };
  }

  it('re-resolves the attachment by name + size and downloads $value', async () => {
    const bytes = new Uint8Array([1, 2, 3]);
    const { src, session } = source((url) => {
      if (url.toString() === `${base}?$select=id,name,size&$top=999`)
        return jsonRes(200, {
          value: [
            { id: 'OTHER', name: 'offer.docx', size: 11 },
            { id: 'RIGHT', name: 'offer.docx', size: 33_630 },
          ],
        });
      if (url.toString() === `${base}/RIGHT/$value`) return bytesRes(bytes);
      return undefined;
    });
    const got = await src.fetchBytes!(
      session,
      doc({ messageId: 'IMMUTABLE-M1', filename: 'offer.docx', sizeBytes: 33_630 }),
    );
    expect([...got!]).toEqual([1, 2, 3]);
  });

  it('never serves a same-named attachment of a different size', async () => {
    const { src, session } = source((url) =>
      url.pathname.endsWith('/attachments')
        ? jsonRes(200, { value: [{ id: 'NEW', name: 'offer.docx', size: 40_000 }] })
        : bytesRes(new Uint8Array([9])),
    );
    await expect(
      src.fetchBytes!(session, doc({ messageId: 'IMMUTABLE-M1', filename: 'offer.docx', sizeBytes: 33_630 })),
    ).resolves.toBeNull();
  });

  it('answers null (terminal) when the message is gone or the attachment no longer exists', async () => {
    const gone = source(() => jsonRes(404, { error: { code: 'ErrorItemNotFound' } }));
    await expect(
      gone.src.fetchBytes!(gone.session, doc({ messageId: 'IMMUTABLE-M1', filename: 'a.pdf', sizeBytes: 1 })),
    ).resolves.toBeNull();
    const missing = source((url) =>
      url.pathname.endsWith('/attachments') ? jsonRes(200, { value: [] }) : undefined,
    );
    await expect(
      missing.src.fetchBytes!(missing.session, doc({ messageId: 'IMMUTABLE-M1', filename: 'a.pdf', sizeBytes: 1 })),
    ).resolves.toBeNull();
  });

  it('lets an auth failure throw (transient — core defers it), not answer null', async () => {
    const { src, session } = source(() => jsonRes(401, { error: { code: 'InvalidAuthenticationToken' } }));
    await expect(
      src.fetchBytes!(session, doc({ messageId: 'IMMUTABLE-M1', filename: 'a.pdf', sizeBytes: 1 })),
    ).rejects.toThrow(/401/);
  });
});
