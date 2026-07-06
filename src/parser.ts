/**
 * Microsoft Graph mail message shapes and parsing, ported verbatim from
 * legacy alpha-cent `src/main/connectors/ms365/parser.ts` (itself sharing
 * `ParsedEmail`/`ParsedAttachment` with `email-shared/types.ts`, inlined here
 * — this connector, like the gmail v2 port, is self-contained). Pure — no
 * network, no I/O.
 *
 * NOTE on attachments: `collectAttachments` is ported faithfully and unit
 * tested, but is DEAD CODE END-TO-END in this connector, same as in legacy —
 * see backfill.ts's module doc for why (the Graph `$select` that fetches
 * conversation messages never requests the `attachments` field, so
 * `msg.attachments` is always `undefined` in practice and this always
 * resolves to `[]`).
 */

export interface GraphEmailAddress {
  address?: string;
  name?: string;
}
export interface GraphRecipient {
  emailAddress?: GraphEmailAddress;
}
export interface GraphAttachmentSummary {
  id?: string;
  '@odata.type'?: string;
  name?: string;
  contentType?: string;
  size?: number;
  contentBytes?: string; // present only when fetched via /attachments
}
export interface GraphInternetHeader {
  name?: string;
  value?: string;
}
export interface GraphMessage {
  id?: string;
  conversationId?: string;
  internetMessageId?: string;
  subject?: string;
  from?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  bccRecipients?: GraphRecipient[];
  receivedDateTime?: string;
  body?: { contentType?: 'text' | 'html'; content?: string };
  hasAttachments?: boolean;
  internetMessageHeaders?: GraphInternetHeader[];
  parentFolderId?: string;
  isDraft?: boolean;
  attachments?: GraphAttachmentSummary[];
}

export interface ParsedAttachment {
  messageId: string;
  partId: string;
  attachmentId: string;
  filename: string;
  mimeType: string;
  sizeBytes: number;
}

export interface ParsedEmail {
  messageId: string;
  threadId: string;
  from: string;
  to: string[];
  cc: string[];
  date: Date;
  subject: string;
  body: string;
  htmlBody: string | null;
  headers: Record<string, string>;
  attachments: ParsedAttachment[];
}

export function parseGraphMessage(msg: GraphMessage): ParsedEmail {
  const headers = collectHeaders(msg.internetMessageHeaders ?? []);
  const isHtml = (msg.body?.contentType ?? 'text') === 'html';
  const content = msg.body?.content ?? '';
  const body = isHtml ? '' : content;
  const htmlBody = isHtml ? content : null;
  return {
    messageId: msg.internetMessageId ?? msg.id ?? '',
    threadId: msg.conversationId ?? '',
    from: formatRecipient(msg.from),
    to: (msg.toRecipients ?? []).map(formatRecipient).filter(Boolean),
    cc: (msg.ccRecipients ?? []).map(formatRecipient).filter(Boolean),
    date: parseDate(msg.receivedDateTime),
    subject: msg.subject ?? '',
    body,
    htmlBody,
    headers,
    attachments: collectAttachments(msg.id ?? '', msg.attachments ?? []),
  };
}

function collectHeaders(hs: GraphInternetHeader[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const h of hs) {
    const k = (h.name ?? '').toLowerCase();
    if (k && !(k in out)) out[k] = h.value ?? '';
  }
  return out;
}

function formatRecipient(r?: GraphRecipient): string {
  const addr = r?.emailAddress?.address ?? '';
  const name = r?.emailAddress?.name ?? '';
  if (!addr) return '';
  if (name && name !== addr) return `${name} <${addr}>`;
  return addr;
}

function parseDate(iso: string | undefined): Date {
  if (!iso) return new Date(0);
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? new Date(0) : d;
}

function collectAttachments(
  graphMessageId: string,
  atts: GraphAttachmentSummary[],
): ParsedAttachment[] {
  const out: ParsedAttachment[] = [];
  for (const a of atts) {
    // Only fileAttachment yields raw bytes via contentBytes. itemAttachment
    // (nested emails) and referenceAttachment (cloud links) are skipped — the
    // first would need recursive ingest, the second has no bytes to fetch.
    if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') continue;
    if (!a.id) continue;
    out.push({
      messageId: graphMessageId,
      partId: '',
      attachmentId: a.id,
      filename: a.name ?? '',
      mimeType: a.contentType ?? 'application/octet-stream',
      sizeBytes: a.size ?? 0,
    });
  }
  return out;
}
