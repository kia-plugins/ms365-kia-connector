/**
 * Reply targets for a conversation — what `metadata.outbound` stores so
 * kiagent-core's `draft_reply` can address a reply WITHOUT the model ever
 * supplying an address. Pure: computed in toDocument from the fetched
 * messages and the account's own address.
 *
 * Mirrors kiagent-core's `resolve-gmail.ts` so a reply behaves the same from
 * Gmail and from Outlook:
 *  - reply: to the last message whose From is not you — its Reply-To when
 *    that names someone else, else its From. When every message is yours,
 *    to the last message's recipients.
 *  - reply-all: the last message's Reply-To (else From) + To, and its Cc —
 *    never you, never twice.
 * Drafts are skipped before anything is chosen: an unsent draft must not
 * pick, or add, a recipient.
 *
 * `ref.messageId` is the IMMUTABLE Graph id of the message replied to (the
 * sender calls `/me/messages/{id}/reply`), never the RFC Message-ID.
 */
import type { GraphMessage, GraphRecipient } from './parser';

export interface ReplyVariant {
  ref: { messageId: string };
  display: string;
  to: string[];
  cc: string[];
}

export interface Ms365Outbound extends ReplyVariant {
  replyAll?: ReplyVariant;
}

const ADDRESS_RX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

const addressOf = (r: GraphRecipient | undefined): string | null => {
  const a = r?.emailAddress?.address?.trim();
  return a && ADDRESS_RX.test(a) ? a : null;
};

const addressesOf = (rs: GraphRecipient[] | undefined): string[] =>
  (rs ?? []).map(addressOf).filter((a): a is string => a !== null);

/** Case-insensitive de-dupe, keeping first spelling; drops `self`. */
function distinct(addrs: string[], self: Set<string>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const a of addrs) {
    const k = a.toLowerCase();
    if (self.has(k) || seen.has(k)) continue;
    seen.add(k);
    out.push(a);
  }
  return out;
}

const display = (to: string[], cc: string[]): string =>
  cc.length ? `${to.join(', ')}; cc ${cc.join(', ')}` : to.join(', ');

export function replyTargets(
  messages: GraphMessage[],
  selfAddress: string | undefined,
): Ms365Outbound | undefined {
  const self = new Set(selfAddress ? [selfAddress.trim().toLowerCase()] : []);
  const isSelf = (m: GraphMessage) => {
    const from = addressOf(m.from);
    return from !== null && self.has(from.toLowerCase());
  };
  const live = messages.filter((m) => !m.isDraft && m.id);
  const last = live[live.length - 1];
  if (!last) return undefined;

  // Reply-To wins — but Graph's "no Reply-To" is an EMPTY list, and a
  // Reply-To pointing back at you never makes you the recipient.
  const primary = (m: GraphMessage): string[] => {
    const replyTo = distinct(addressesOf(m.replyTo), self);
    if (replyTo.length) return replyTo;
    const from = addressOf(m.from);
    return from ? [from] : [];
  };

  let reply: ReplyVariant | undefined;
  const target = [...live].reverse().find((m) => !isSelf(m));
  if (target) {
    const to = distinct(primary(target), self);
    if (to.length)
      reply = { ref: { messageId: target.id! }, display: display(to, []), to, cc: [] };
  } else {
    const to = distinct(addressesOf(last.toRecipients), self);
    if (to.length)
      reply = { ref: { messageId: last.id! }, display: display(to, []), to, cc: [] };
  }

  const allTo = distinct([...primary(last), ...addressesOf(last.toRecipients)], self);
  const allCc = distinct(addressesOf(last.ccRecipients), self).filter(
    (a) => !allTo.some((t) => t.toLowerCase() === a.toLowerCase()),
  );
  const replyAll: ReplyVariant | undefined = allTo.length
    ? { ref: { messageId: last.id! }, display: display(allTo, allCc), to: allTo, cc: allCc }
    : undefined;

  if (!reply) return undefined;
  return replyAll ? { ...reply, replyAll } : reply;
}
