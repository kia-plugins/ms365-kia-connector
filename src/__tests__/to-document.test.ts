/**
 * Ported from legacy `src/__tests__/ms365-thread-builder.test.ts`
 * (buildMs365Thread), reshaped as a PURE toDocument test (no DB, no
 * converter — see src/to-document.ts's module doc for what changed).
 */
import { buildThreadUrl, EMAIL_THREAD_DOCUMENT_TYPE, toDocument } from '../to-document';
import { graphMsg } from '../testing/harness';

describe('toDocument', () => {
  it('builds an email.thread document with markdown, participants, and metadata', () => {
    const doc = toDocument({
      conversationId: 'C1',
      tenantKind: 'personal',
      messages: [graphMsg()],
    })!;

    expect(doc.externalId).toBe('C1');
    expect(doc.type).toBe(EMAIL_THREAD_DOCUMENT_TYPE);
    expect(doc.title).toBe('hello');
    expect(doc.markdown).toMatch(/^# hello/);
    expect(doc.markdown).toContain('## 1 — A <a@x.com> · 2026-05-20 10:00');
    expect(doc.url).toBe(buildThreadUrl('personal', 'C1'));
    expect(doc.metadata).toMatchObject({
      ms365ConversationId: 'C1',
      from: 'A <a@x.com>',
      messageCount: 1,
      tenantKind: 'personal',
    });
    expect(doc.createdAt).toBe('2026-05-20T10:00:00.000Z');
  });

  it('returns null for a zero-message item', () => {
    expect(toDocument({ conversationId: 'C1', tenantKind: 'personal', messages: [] })).toBeNull();
  });

  it.each([
    ['empty-string', ''],
    ['whitespace-only', '   '],
  ])('falls back to (no subject) for a %s subject', (_label, subject) => {
    const doc = toDocument({
      conversationId: 'C1',
      tenantKind: 'personal',
      messages: [graphMsg({ subject })],
    })!;
    expect(doc.title).toBe('(no subject)');
    expect(doc.markdown).toMatch(/^# \(no subject\)/);
  });

  it('stamps createdAt from the LAST message (recency ordering — see module doc)', () => {
    const doc = toDocument({
      conversationId: 'C1',
      tenantKind: 'personal',
      messages: [
        graphMsg({ id: 'm1', receivedDateTime: '2026-05-20T10:00:00Z', body: { contentType: 'text', content: 'first' } }),
        graphMsg({ id: 'm2', receivedDateTime: '2026-05-21T10:00:00Z', body: { contentType: 'text', content: 'second' } }),
      ],
    })!;
    expect(doc.createdAt).toBe('2026-05-21T10:00:00.000Z');
    expect(doc.metadata.firstMessageAt).toBe('2026-05-20T10:00:00.000Z');
    expect(doc.metadata.lastMessageAt).toBe('2026-05-21T10:00:00.000Z');
  });

  it('dedupes participants across from/to/cc of every message', () => {
    const doc = toDocument({
      conversationId: 'C1',
      tenantKind: 'personal',
      messages: [
        graphMsg({
          from: { emailAddress: { address: 'a@x.com' } },
          toRecipients: [{ emailAddress: { address: 'b@x.com' } }],
        }),
        graphMsg({
          id: 'm2',
          from: { emailAddress: { address: 'b@x.com' } },
          toRecipients: [{ emailAddress: { address: 'a@x.com' } }],
        }),
      ],
    })!;
    expect(doc.metadata.participants).toEqual(['a@x.com', 'b@x.com']);
  });

  it('does NOT filter automated/list-unsubscribe threads (dropped vs legacy — see module doc)', () => {
    const doc = toDocument({
      conversationId: 'C1',
      tenantKind: 'personal',
      messages: [
        graphMsg({
          internetMessageHeaders: [{ name: 'List-Unsubscribe', value: '<https://x/unsub>' }],
        }),
      ],
    });
    expect(doc).not.toBeNull();
  });
});

describe('buildThreadUrl', () => {
  it('uses outlook.office.com for a work tenant', () => {
    const url = buildThreadUrl('work', 'C1');
    expect(url).toMatch(/^https:\/\/outlook\.office\.com\/mail\/inbox\/id\//);
  });

  it('uses outlook.live.com for a personal tenant', () => {
    const url = buildThreadUrl('personal', 'C1');
    expect(url).toMatch(/^https:\/\/outlook\.live\.com\/mail\/0\/inbox\/id\//);
  });

  it('base64url-encodes the conversationId as the trailing id segment', () => {
    const url = buildThreadUrl('work', 'C1');
    const id = url.split('/').pop()!;
    expect(Buffer.from(id, 'base64url').toString('utf-8')).toBe('C1');
  });
});
