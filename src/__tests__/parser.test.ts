/**
 * Ported from the legacy v1 repo's `src/__tests__/ms365-parser.test.ts` —
 * parseGraphMessage is unchanged from v1 (see src/parser.ts's module doc).
 */
import { parseGraphMessage } from '../parser';

describe('ms365 parseGraphMessage', () => {
  it('parses a minimal text message', () => {
    const parsed = parseGraphMessage({
      id: 'AAA',
      conversationId: 'C1',
      internetMessageId: '<a@x>',
      subject: 'hello',
      from: { emailAddress: { address: 'alice@x.com', name: 'Alice' } },
      toRecipients: [{ emailAddress: { address: 'bob@y.com', name: 'Bob' } }],
      ccRecipients: [],
      receivedDateTime: '2026-05-20T10:00:00Z',
      body: { contentType: 'text', content: 'hi there' },
      hasAttachments: false,
      internetMessageHeaders: [{ name: 'Auto-Submitted', value: 'no' }],
      parentFolderId: 'inbox',
    });
    expect(parsed.threadId).toBe('C1');
    expect(parsed.from).toBe('Alice <alice@x.com>');
    expect(parsed.to).toEqual(['Bob <bob@y.com>']);
    expect(parsed.subject).toBe('hello');
    expect(parsed.body).toBe('hi there');
    expect(parsed.htmlBody).toBeNull();
    expect(parsed.date.toISOString()).toBe('2026-05-20T10:00:00.000Z');
    expect(parsed.headers['auto-submitted']).toBe('no');
    expect(parsed.attachments).toEqual([]);
  });

  it('keeps htmlBody when contentType is html and body empty', () => {
    const parsed = parseGraphMessage({
      id: 'AAA',
      conversationId: 'C1',
      internetMessageId: '<a@x>',
      subject: 's',
      from: { emailAddress: { address: 'a@x.com' } },
      toRecipients: [],
      ccRecipients: [],
      receivedDateTime: '2026-05-20T10:00:00Z',
      body: { contentType: 'html', content: '<p>hi</p>' },
      hasAttachments: false,
      internetMessageHeaders: [],
      parentFolderId: 'inbox',
    });
    expect(parsed.body).toBe('');
    expect(parsed.htmlBody).toBe('<p>hi</p>');
  });

  it('falls back to epoch when receivedDateTime is missing', () => {
    const parsed = parseGraphMessage({
      id: 'AAA',
      conversationId: 'C1',
      internetMessageId: '<a@x>',
      subject: 's',
      from: { emailAddress: { address: 'a@x.com' } },
      toRecipients: [],
      ccRecipients: [],
      receivedDateTime: undefined,
      body: { contentType: 'text', content: '' },
      hasAttachments: false,
      internetMessageHeaders: [],
      parentFolderId: 'inbox',
    });
    expect(parsed.date.getTime()).toBe(0);
  });

  it('formats recipients with name, and bare address when name equals address', () => {
    const parsed = parseGraphMessage({
      from: { emailAddress: { address: 'a@x.com', name: 'a@x.com' } },
      toRecipients: [{ emailAddress: { address: 'b@x.com' } }],
    });
    expect(parsed.from).toBe('a@x.com');
    expect(parsed.to).toEqual(['b@x.com']);
  });

  it('drops non-fileAttachment entries and entries without an id (dead code today — see module doc)', () => {
    const parsed = parseGraphMessage({
      id: 'M1',
      attachments: [
        { '@odata.type': '#microsoft.graph.itemAttachment', id: 'A0', name: 'nested.eml' },
        { '@odata.type': '#microsoft.graph.fileAttachment', name: 'no-id.pdf' },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'A1',
          name: 'doc.pdf',
          contentType: 'application/pdf',
          size: 42,
        },
      ],
    });
    expect(parsed.attachments).toEqual([
      {
        messageId: 'M1',
        partId: '',
        attachmentId: 'A1',
        filename: 'doc.pdf',
        mimeType: 'application/pdf',
        sizeBytes: 42,
      },
    ]);
  });
});
