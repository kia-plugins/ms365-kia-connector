import type { DocumentInput } from '@kiagent/connector-sdk';
import { parseGraphMessage, type GraphMessage } from './parser';

/** Shared with the gmail v2 builtin's document-type convention — this
 *  connector emits the SAME generic `email.thread` type (verified against
 *  `kiagent-core/src/main/sources/gmail/to-document.ts`) so Outlook mail
 *  renders identically to Gmail mail in the app (same search/type filters,
 *  same rendering surface). See report for the full doc-type-naming
 *  writeup. */
export const EMAIL_THREAD_DOCUMENT_TYPE = 'email.thread';

/**
 * What `pull()` fetches per conversation and hands to `toDocument`. This is
 * the `Item` half of `Source<Cursor, Item>` — plain, already-fetched, fully
 * self-describing, so `toDocument` can stay PURE (no session, no network).
 * `tenantKind` is stamped on by `pull()` from `session.account.config` (set
 * once at connect time — see source.ts) so the deep link can pick the right
 * outlook.office.com/outlook.live.com host without `toDocument` needing
 * account context.
 */
export interface Ms365ThreadItem {
  /** Graph conversationId — the externalId scheme (unchanged from legacy). */
  conversationId: string;
  /** Raw messages from `/me/messages?$filter=conversationId eq '…'`,
   *  oldest first. */
  messages: GraphMessage[];
  tenantKind: 'work' | 'personal';
}

/**
 * `work` tenants use the Outlook web app hosted at outlook.office.com,
 * `personal` (MSA) tenants use outlook.live.com — ported EXACTLY from
 * legacy's `ms365/index.ts` `buildSourceUrl` (base64url of the item's own
 * id). Legacy's `thread-builder.ts` had a SEPARATE, more elaborate
 * `buildConversationUrl` that based the id on the first message's
 * `messageId` instead of the conversationId; this port intentionally uses
 * the simpler, pure, dependency-free `buildSourceUrl` form the task brief
 * names explicitly — see report.
 */
export function buildThreadUrl(
  tenantKind: 'work' | 'personal',
  conversationId: string,
): string {
  const id = Buffer.from(conversationId, 'utf-8').toString('base64url');
  return tenantKind === 'work'
    ? `https://outlook.office.com/mail/inbox/id/${id}`
    : `https://outlook.live.com/mail/0/inbox/id/${id}`;
}

/** PURE conversation → DocumentInput mapping. Returns null for a
 *  conversation with zero messages (mirrors legacy's "empty conversation"
 *  skip — pull() never actually emits such an item; see backfill.ts /
 *  delta.ts, which route a zero-message conversation to a deletion instead).
 *
 *  DROPPED vs legacy: legacy's `isAutomatedThread` filter (Auto-Submitted /
 *  Precedence / List-* headers / system-sender local-parts / empty
 *  Return-Path / DSN multipart-report) is NOT ported — matching the gmail
 *  v2 builtin's precedent (`to-document.ts` there has no such filter
 *  either). Every conversation is now indexed regardless of these headers.
 *  DROPPED vs legacy: no attachment sub-documents (see graph-api.ts's
 *  CONV_SELECT comment — legacy's own attachment wiring was unreachable).
 *  DROPPED vs legacy: HTML-only message bodies render as an empty body
 *  (legacy called its DB-backed `Converter` to turn `htmlBody` into
 *  markdown at ingest time; this pure v2 mapping has no converter to call,
 *  matching the same simplification already made in the gmail v2 port). */
export function toDocument(item: Ms365ThreadItem): DocumentInput | null {
  if (item.messages.length === 0) return null;

  const parsed = item.messages.map(parseGraphMessage);
  const first = parsed[0];
  const last = parsed[parsed.length - 1];

  const subject = (first.subject?.trim() || '(no subject)') as string;
  const url = buildThreadUrl(item.tenantKind, item.conversationId);

  const sections: string[] = [];
  sections.push(`# ${subject}\n`);
  sections.push(
    `> Thread: ${parsed.length} messages · ${fmt(first.date)} → ${fmt(last.date)}`,
  );
  sections.push(`> Open in Outlook: ${url}\n\n---`);

  let idx = 1;
  for (const m of parsed) {
    sections.push(`## ${idx} — ${m.from} · ${fmt(m.date)}\n\n${m.body}`);
    idx += 1;
  }

  const participants = [
    ...new Set(parsed.flatMap((m) => [m.from, ...m.to, ...m.cc])),
  ];

  return {
    externalId: item.conversationId,
    type: EMAIL_THREAD_DOCUMENT_TYPE,
    title: subject,
    markdown: sections.join('\n'),
    url,
    metadata: {
      ms365ConversationId: item.conversationId,
      from: first.from,
      to: first.to,
      cc: first.cc,
      messageCount: parsed.length,
      participants,
      firstMessageAt: first.date.toISOString(),
      lastMessageAt: last.date.toISOString(),
      tenantKind: item.tenantKind,
      messages: parsed.map((m) => ({
        id: m.messageId,
        from: m.from,
        date: m.date.toISOString(),
        snippet: m.body.slice(0, 200),
      })),
    },
    // Last message date, matching the gmail v2 port's same deliberate
    // deviation from legacy (which stamped created_at from the FIRST
    // message) — see that port's report for rationale (recency ordering).
    createdAt: last.date.toISOString(),
  };
}

function fmt(d: Date): string {
  return d.toISOString().slice(0, 16).replace('T', ' ');
}
