/**
 * Microsoft 365 Sender — replies and new mail through Microsoft Graph.
 * Reachable only from kiagent-core's confirmation-gated send pipeline.
 *
 *  - reply: `POST /me/messages/{id}/reply` on the message the stored reply
 *    target names (reply-target.ts), with EXACTLY the recipients the user
 *    confirmed (`intent.to/cc`, frozen on the draft) — Graph threads it.
 *  - new mail: `POST /me/sendMail`, saved to Sent Items.
 *
 * Both answer 202 with no id. The body is sent as plain text — the same
 * thing the Gmail and IMAP senders send.
 *
 * Failure wording is a cross-repo contract with kiagent-core's
 * error-copy.ts: `reconnect … in Settings` (auth), `rate-limited:` and
 * `not sent:` mark failures that PROVE nothing left; anything else reads
 * as "may have been sent". A network error or 5xx is never retried
 * (GraphClient.post) — Graph may already have accepted the message.
 */
import type {
  HostFor,
  SendIntent,
  SendResult,
  Sender,
  SenderContext,
} from '@kiagent/connector-sdk';
import { GRAPH_BASE } from './graph-api';
import {
  GraphApiError,
  GraphClient,
  isAuthError,
  type GraphClientDeps,
} from './graph-client';

/** Well inside kiagent-core's 60 s sender timeout, which does not cancel
 *  this call — nothing may still be in flight once it has given up. */
export const SEND_DEADLINE_MS = 40_000;

const recipients = (addrs: string[] | undefined) =>
  (addrs ?? []).map((address) => ({ emailAddress: { address } }));

export function createMs365Sender(
  host: HostFor<'net'>,
  clock?: Pick<GraphClientDeps, 'sleep' | 'random'>,
): Sender {
  return {
    async send(intent: SendIntent, ctx?: SenderContext): Promise<SendResult> {
      // Refreshed host-side at send time — this process has no vault.
      const token = ctx?.credentials?.accessToken;
      if (!token)
        throw new Error('no Microsoft 365 credentials — reconnect the account in Settings');
      const to = intent.to ?? [];
      if (to.length === 0) throw new Error('not sent: this draft has no recipient');
      const body = { contentType: 'Text', content: intent.bodyMarkdown };

      let url: string;
      let payload: unknown;
      if (intent.kind === 'reply') {
        const messageId = (intent.outboundRef as { messageId?: unknown } | undefined)?.messageId;
        if (typeof messageId !== 'string' || messageId.length === 0)
          throw new Error('not sent: this draft has no message to reply to');
        url = `${GRAPH_BASE}/me/messages/${encodeURIComponent(messageId)}/reply`;
        payload = {
          message: { toRecipients: recipients(to), ccRecipients: recipients(intent.cc), body },
        };
      } else {
        url = `${GRAPH_BASE}/me/sendMail`;
        payload = {
          message: {
            subject: intent.subject ?? '',
            body,
            toRecipients: recipients(to),
            ccRecipients: recipients(intent.cc),
          },
          saveToSentItems: true,
        };
      }

      const client = new GraphClient({
        fetch: host.net.fetch,
        getToken: async () => token,
        ...clock,
      });
      try {
        await client.post(url, payload, { deadlineMs: SEND_DEADLINE_MS });
        return {};
      } catch (e) {
        throw sendError(e);
      }
    },
  };
}

/** Graph's answer → the outbound layer's wording. Only statuses that prove
 *  Graph REFUSED the request become a "nothing was sent" marker. */
function sendError(e: unknown): Error {
  if (isAuthError(e))
    return new Error('your Microsoft 365 sign-in no longer works — reconnect the account in Settings');
  if (!(e instanceof GraphApiError)) return e instanceof Error ? e : new Error(String(e));
  const code = e.code ?? '';
  if (e.status === 403 && /AccessDenied|Authorization_RequestDenied/i.test(code))
    return new Error(
      'this Microsoft 365 account was connected before sending existed — reconnect it in Settings to grant send permission',
    );
  if (e.status === 404 && /ErrorItemNotFound/.test(code))
    return new Error('not sent: the original message no longer exists in the mailbox');
  if (e.status === 429)
    return new Error('rate-limited: Microsoft 365 is throttling sends — nothing was sent');
  if (e.status === 400)
    return new Error(`not sent: Microsoft 365 rejected the message${code ? ` (${code})` : ''}`);
  return e;
}
