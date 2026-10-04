/**
 * Owner-channel listing (Email Kia): every message addressed To one address.
 *
 *  - Sent leg: the account's own requests, from Sent Items, read in full
 *    (headers + plain-text body) — only these can become assistant turns.
 *  - Audit leg: everything else addressed there, mailbox-wide (Junk and
 *    Deleted Items included, drafts excluded), as metadata only — so the
 *    assistant can log it while its body never leaves Graph.
 *
 * Addresses are lowercased: the assistant's gate compares them against a
 * lowercased owner address.
 */
import type { AddressedMail, AddressedQuery } from '@kiagent/connector-sdk';
import { GRAPH_BASE } from './graph-api';
import type { GraphClient } from './graph-client';
import type { GraphInternetHeader, GraphRecipient } from './parser';

const TEXT_CAP = 32 * 1024;
const PAGE = '100';
const SENT_SELECT =
  'id,internetMessageId,toRecipients,ccRecipients,sentDateTime,from,sender,subject,conversationId';
const AUDIT_SELECT =
  'id,internetMessageId,parentFolderId,isDraft,from,sender,toRecipients,ccRecipients,subject,receivedDateTime';

interface ListedMessage {
  id?: string;
  internetMessageId?: string;
  parentFolderId?: string;
  isDraft?: boolean;
  from?: GraphRecipient;
  sender?: GraphRecipient;
  toRecipients?: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  subject?: string;
  sentDateTime?: string;
  receivedDateTime?: string;
  conversationId?: string;
}

interface FullMessage {
  internetMessageHeaders?: GraphInternetHeader[];
  body?: { contentType?: string; content?: string };
  conversationId?: string;
}

const iso = (ms: number) => new Date(ms).toISOString();
const addr = (r?: GraphRecipient) => (r?.emailAddress?.address ?? '').trim().toLowerCase();
const addrs = (rs?: GraphRecipient[]) => (rs ?? []).map(addr).filter(Boolean);
const ids = (v = '') => v.match(/<[^<>\s]+>/g) ?? [];

function headersOf(hs: GraphInternetHeader[] = []): Record<string, string> {
  const out: Record<string, string> = {};
  for (const h of hs) {
    const k = (h.name ?? '').toLowerCase();
    if (k && !(k in out)) out[k] = h.value ?? '';
  }
  return out;
}

/** kiagent-core's `isAutomatedMessage` header checks (imap/filter.ts). */
function isAutomated(h: Record<string, string>): boolean {
  const autoSub = h['auto-submitted'];
  if (autoSub && autoSub.toLowerCase() !== 'no') return true;
  if (['bulk', 'list', 'junk', 'auto_reply'].includes((h.precedence ?? '').toLowerCase())) return true;
  if (['list-id', 'list-unsubscribe', 'list-post'].some((k) => h[k])) return true;
  const suppress = (h['x-auto-response-suppress'] ?? '').toLowerCase();
  if (['all', 'autoreply', 'oof', 'dr'].some((t) => suppress.includes(t))) return true;
  return !!(h['x-autoreply'] || h['x-autorespond']);
}

async function* pages(client: GraphClient, path: string, filter: string, select: string) {
  const first = new URL(`${GRAPH_BASE}${path}`);
  first.searchParams.set('$filter', filter);
  first.searchParams.set('$select', select);
  first.searchParams.set('$top', PAGE);
  let url: string | undefined = first.toString();
  while (url) {
    // eslint-disable-next-line no-await-in-loop
    const page: { value?: ListedMessage[]; '@odata.nextLink'?: string } = await client.request(url);
    yield* page.value ?? [];
    url = page['@odata.nextLink'];
  }
}

function base(m: ListedMessage, date: string | undefined) {
  const from = addr(m.from);
  const sender = addr(m.sender);
  return {
    providerId: m.id ?? '',
    from,
    ...(sender && sender !== from ? { sender } : {}),
    subject: m.subject ?? '',
    date: Date.parse(date ?? '') || 0,
    to: addrs(m.toRecipients),
    cc: addrs(m.ccRecipients),
  };
}

export async function listAddressedTo(
  client: GraphClient,
  q: AddressedQuery,
  sentFolderId: string,
): Promise<AddressedMail[]> {
  const target = q.toAddress.toLowerCase();
  const toTarget = (m: ListedMessage) => addrs(m.toRecipients).includes(target);
  const out: AddressedMail[] = [];
  const seen = new Set<string>();

  const sent = pages(client, '/me/mailFolders/sentitems/messages', `sentDateTime ge ${iso(q.since)}`, SENT_SELECT);
  for await (const m of sent) {
    if (!m.id || !toTarget(m)) continue;
    const full: FullMessage = await client.request(
      `${GRAPH_BASE}/me/messages/${encodeURIComponent(m.id)}?$select=internetMessageHeaders,body,conversationId`,
      { extraHeaders: { prefer: 'outlook.body-content-type="text"' } },
    );
    const h = headersOf(full.internetMessageHeaders);
    const messageId = m.internetMessageId ?? ids(h['message-id'])[0] ?? '';
    if (messageId) seen.add(messageId);
    const inReplyTo = ids(h['in-reply-to'])[0];
    const thread = full.conversationId ?? m.conversationId;
    out.push({
      ...base(m, m.sentDateTime),
      messageId,
      ...(inReplyTo ? { inReplyTo } : {}),
      references: ids(h.references),
      ...(thread ? { providerThreadId: thread } : {}),
      sent: true,
      text: (full.body?.content ?? '').slice(0, TEXT_CAP),
      ownChannel: 'x-kia-channel' in h,
      automated: isAutomated(h),
    } as AddressedMail);
  }

  const audit = pages(
    client,
    '/me/messages',
    `receivedDateTime ge ${iso(q.auditSince ?? q.since)}`,
    AUDIT_SELECT,
  );
  for await (const m of audit) {
    if (!m.id || m.isDraft || m.parentFolderId === sentFolderId || !toTarget(m)) continue;
    // The inbox copy of a request you sent to yourself.
    if (m.internetMessageId && seen.has(m.internetMessageId)) continue;
    out.push({
      ...base(m, m.receivedDateTime),
      messageId: m.internetMessageId ?? '',
      references: [],
      sent: false,
      text: '',
      ownChannel: false,
      automated: false,
    } as AddressedMail);
  }
  return out;
}
