/**
 * The Outlook side of an Outbox draft (alpha-cent
 * docs/superpowers/specs/2026-10-20-outbox-draft-sync-design.md §5.6): a real
 * Outlook draft kept in step with the Kia row, so the user can finish it in
 * Outlook and Kia sees what they changed.
 *
 *  - create: a reply is `createReply` on the message the reply target names,
 *    then a PATCH to exactly the confirmed recipients and Kia's text; new
 *    mail is `POST /me/messages`. Files go on after (upload session above
 *    3 MB). If anything after the create fails, the half-made draft is
 *    deleted so no stray draft is left in Outlook.
 *  - update: PATCHes only the fields the patch names — the HTML body,
 *    signature and files the user added in Outlook stay unless the patch
 *    replaces them.
 *  - read: `isDraft: false` → sent; in Deleted Items → deleted (deleting a
 *    draft in Outlook moves it, and the immutable id survives the move);
 *    gone (404) → gone.
 *  - send: `POST …/send` — the draft exactly as Outlook holds it. Never
 *    retried after a network error or 5xx (GraphClient.post).
 *
 * Every call is bounded by one deadline; nothing it started is still in
 * flight once it has thrown.
 */
import type {
  DraftContent,
  DraftMirror,
  DraftPatch,
  DraftReadResult,
  DraftRef,
  HostFor,
  SendIntent,
  SendResult,
  SenderContext,
} from '@kiagent/connector-sdk';
import { GRAPH_BASE } from './graph-api';
import {
  GraphApiError,
  GraphClient,
  isAuthError,
  type GraphClientDeps,
} from './graph-client';
import type { GraphAttachment, GraphMessage, GraphRecipient } from './parser';
import { SEND_DEADLINE_MS, sendError } from './sender';

/** Matches kiagent-core's error-copy `reconnect .* in Settings` contract. */
export const DRAFT_RECONNECT = 'Reconnect this account in Settings to sync drafts';

/** create / read / remove: inside kiagent-core's 20 s mirror deadline. */
export const DRAFT_DEADLINE_MS = 18_000;
/** update: core awaits it to the end (its RPC allows 60 s). */
export const DRAFT_UPDATE_DEADLINE_MS = 40_000;
/** Graph takes a file in one POST below 3 MB; larger ones need a session. */
export const SIMPLE_ATTACHMENT_MAX_BYTES = 3_000_000;
/** Upload-session chunks must be multiples of 320 KiB. */
export const UPLOAD_CHUNK_BYTES = 320 * 1024 * 10;

type FileToAdd = NonNullable<SendIntent['attachments']>[number];

const recipients = (addrs: string[] | undefined) =>
  (addrs ?? []).map((address) => ({ emailAddress: { address } }));

const addressesOf = (list: GraphRecipient[] | undefined) =>
  (list ?? []).map((r) => r.emailAddress?.address).filter((a): a is string => !!a);

const msgUrl = (id: string) => `${GRAPH_BASE}/me/messages/${encodeURIComponent(id)}`;

const READ_SELECT =
  'isDraft,parentFolderId,internetMessageId,conversationId,subject,' +
  'toRecipients,ccRecipients,bccRecipients,body';

const isItemGone = (e: unknown) =>
  e instanceof GraphApiError && e.status === 404 && e.code === 'ErrorItemNotFound';

/** A draft write's failure in the wording kiagent-core shows. */
function draftError(e: unknown): Error {
  if (isAuthError(e))
    return new Error('your Microsoft 365 sign-in no longer works — reconnect the account in Settings');
  if (
    e instanceof GraphApiError &&
    e.status === 403 &&
    /AccessDenied|Authorization_RequestDenied/i.test(e.code ?? '')
  )
    return new Error(DRAFT_RECONNECT);
  if (isItemGone(e)) return new Error('the Outlook draft or the message it answers no longer exists');
  return e instanceof Error ? e : new Error(String(e));
}

export function createMs365Drafts(
  host: HostFor<'net'>,
  clock?: Pick<GraphClientDeps, 'sleep' | 'random'> & { now?: () => number },
): DraftMirror {
  const now = clock?.now ?? Date.now;
  /** Deleted Items' id per account — a well-known folder never moves.
   *  Only a found id is kept; a miss is asked again next time. */
  const deletedItems = new Map<string, string>();

  const clientFor = (ctx?: SenderContext) => {
    // Refreshed host-side per call — this process has no vault.
    const token = ctx?.credentials?.accessToken;
    if (!token) throw new Error('no Microsoft 365 credentials — reconnect the account in Settings');
    return new GraphClient({ fetch: host.net.fetch, getToken: async () => token, ...clock });
  };

  async function addFile(client: GraphClient, id: string, a: FileToAdd, deadline: number) {
    const bytes = Buffer.from(a.contentBase64, 'base64');
    if (bytes.byteLength < SIMPLE_ATTACHMENT_MAX_BYTES) {
      await client.write('POST', `${msgUrl(id)}/attachments`, {
        body: {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: a.filename,
          contentType: a.mime,
          contentBytes: a.contentBase64,
        },
        deadline,
        idempotent: false,
        now,
      });
      return;
    }
    const session = await client.write<{ uploadUrl?: string }>(
      'POST',
      `${msgUrl(id)}/attachments/createUploadSession`,
      {
        body: {
          AttachmentItem: {
            attachmentType: 'file',
            name: a.filename,
            size: bytes.byteLength,
            contentType: a.mime,
          },
        },
        deadline,
        idempotent: false,
        now,
      },
    );
    if (!session?.uploadUrl) throw new Error(`Outlook gave no upload URL for "${a.filename}"`);
    for (let start = 0; start < bytes.byteLength; start += UPLOAD_CHUNK_BYTES) {
      const end = Math.min(start + UPLOAD_CHUNK_BYTES, bytes.byteLength);
      // Pre-authenticated URL: a bearer token is refused. Re-sending the
      // same range is safe.
      // eslint-disable-next-line no-await-in-loop
      await client.write('PUT', session.uploadUrl, {
        raw: bytes.subarray(start, end),
        headers: {
          'content-type': 'application/octet-stream',
          'content-range': `bytes ${start}-${end - 1}/${bytes.byteLength}`,
        },
        auth: false,
        deadline,
        idempotent: true,
        now,
      });
    }
  }

  async function addFiles(client: GraphClient, id: string, files: FileToAdd[], deadline: number) {
    for (const a of files) {
      // eslint-disable-next-line no-await-in-loop
      await addFile(client, id, a, deadline);
    }
  }

  async function removeQuietly(client: GraphClient, id: string, deadline: number) {
    try {
      await client.write('DELETE', msgUrl(id), { deadline, idempotent: true, now });
    } catch {
      // best effort: the create already failed and that error is the one to show
    }
  }

  /** Under the read's own deadline, like every call here. */
  async function deletedItemsId(accountId: string, client: GraphClient, deadline: number) {
    const known = deletedItems.get(accountId);
    if (known) return known;
    const f = await client.write<{ id?: string }>(
      'GET',
      `${GRAPH_BASE}/me/mailFolders/deleteditems?$select=id`,
      { deadline, idempotent: true, now },
    );
    if (f?.id) deletedItems.set(accountId, f.id);
    return f?.id;
  }

  return {
    async create(intent, ctx): Promise<DraftRef> {
      const client = clientFor(ctx);
      const deadline = now() + DRAFT_DEADLINE_MS;
      const body = { contentType: 'Text', content: intent.bodyMarkdown };
      let made: GraphMessage | undefined;
      try {
        if (intent.kind === 'reply') {
          const messageId = (intent.outboundRef as { messageId?: unknown } | undefined)?.messageId;
          if (typeof messageId !== 'string' || messageId.length === 0)
            throw new Error('this draft has no message to reply to');
          made = await client.write<GraphMessage>('POST', `${msgUrl(messageId)}/createReply`, {
            body: {},
            deadline,
            idempotent: false,
            now,
          });
          if (!made?.id) throw new Error('Outlook made no draft');
          // createReply addresses the sender only; the confirmed
          // recipients and Kia's text replace what it filled in.
          await client.write('PATCH', msgUrl(made.id), {
            body: {
              toRecipients: recipients(intent.to),
              ccRecipients: recipients(intent.cc),
              bccRecipients: recipients(intent.bcc),
              ...(intent.subject ? { subject: intent.subject } : {}),
              body,
            },
            deadline,
            idempotent: true,
            now,
          });
        } else {
          made = await client.write<GraphMessage>('POST', `${GRAPH_BASE}/me/messages`, {
            body: {
              subject: intent.subject ?? '',
              body,
              toRecipients: recipients(intent.to),
              ccRecipients: recipients(intent.cc),
              bccRecipients: recipients(intent.bcc),
            },
            deadline,
            idempotent: false,
            now,
          });
          if (!made?.id) throw new Error('Outlook made no draft');
        }
        await addFiles(client, made.id, intent.attachments ?? [], deadline);
      } catch (e) {
        if (made?.id) await removeQuietly(client, made.id, now() + 5_000);
        throw draftError(e);
      }
      return {
        accountId: intent.accountId,
        id: made.id,
        ...(made.conversationId ? { threadId: made.conversationId } : {}),
      };
    },

    async update(ref, patch: DraftPatch, ctx): Promise<DraftRef> {
      const client = clientFor(ctx);
      const deadline = now() + DRAFT_UPDATE_DEADLINE_MS;
      const fields = {
        ...(patch.to ? { toRecipients: recipients(patch.to) } : {}),
        ...(patch.cc ? { ccRecipients: recipients(patch.cc) } : {}),
        ...(patch.subject !== undefined ? { subject: patch.subject ?? '' } : {}),
        ...(patch.bodyMarkdown !== undefined
          ? { body: { contentType: 'Text', content: patch.bodyMarkdown } }
          : {}),
      };
      try {
        if (Object.keys(fields).length)
          await client.write('PATCH', msgUrl(ref.id), {
            body: fields,
            deadline,
            idempotent: true,
            now,
          });
        const removed = new Set(patch.removeAttachments ?? []);
        if (removed.size) {
          const list = await client.write<{ value?: GraphAttachment[] }>(
            'GET',
            `${msgUrl(ref.id)}/attachments?$select=id,name,isInline`,
            { deadline, idempotent: true, now },
          );
          for (const a of list?.value ?? []) {
            if (!a.id || a.isInline || !removed.has(a.name ?? '')) continue;
            try {
              // eslint-disable-next-line no-await-in-loop
              await client.write(
                'DELETE',
                `${msgUrl(ref.id)}/attachments/${encodeURIComponent(a.id)}`,
                { deadline, idempotent: true, now },
              );
            } catch (e) {
              if (!isItemGone(e)) throw e;
            }
          }
        }
        await addFiles(client, ref.id, patch.attachments ?? [], deadline);
      } catch (e) {
        throw draftError(e);
      }
      return ref;
    },

    async read(ref, ctx): Promise<DraftReadResult> {
      const client = clientFor(ctx);
      const deadline = now() + DRAFT_DEADLINE_MS;
      let m: GraphMessage | undefined;
      try {
        m = await client.write<GraphMessage>('GET', `${msgUrl(ref.id)}?$select=${READ_SELECT}`, {
          headers: { prefer: 'outlook.body-content-type="text"' },
          deadline,
          idempotent: true,
          now,
        });
      } catch (e) {
        if (isItemGone(e)) return { state: 'gone' };
        throw draftError(e);
      }
      if (!m) return { state: 'gone' };
      if (m.isDraft === false)
        return {
          state: 'sent',
          ...(m.internetMessageId ? { externalMessageId: m.internetMessageId } : {}),
        };
      if (m.parentFolderId && m.parentFolderId === (await deletedItemsId(ref.accountId, client, deadline)))
        return { state: 'deleted' };
      const files = await client.write<{ value?: GraphAttachment[] }>(
        'GET',
        `${msgUrl(ref.id)}/attachments?$select=name,contentType,size,isInline`,
        { deadline, idempotent: true, now },
      );
      const content: DraftContent = {
        to: addressesOf(m.toRecipients),
        cc: addressesOf(m.ccRecipients),
        bcc: addressesOf(m.bccRecipients),
        subject: m.subject ?? null,
        bodyText: (m.body?.content ?? '').replace(/\r\n/g, '\n'),
        attachments: (files?.value ?? [])
          .filter((a) => !a.isInline && a.name)
          .map((a) => ({
            filename: a.name!,
            mime: a.contentType ?? 'application/octet-stream',
            size: a.size ?? 0,
          })),
      };
      return {
        state: 'draft',
        ref: { ...ref, ...(m.conversationId ? { threadId: m.conversationId } : {}) },
        content,
      };
    },

    async remove(ref, ctx): Promise<void> {
      const client = clientFor(ctx);
      try {
        await client.write('DELETE', msgUrl(ref.id), {
          deadline: now() + DRAFT_DEADLINE_MS,
          idempotent: true,
          now,
        });
      } catch (e) {
        if (!isItemGone(e)) throw draftError(e);
      }
    },

    async send(ref, ctx): Promise<SendResult> {
      const client = clientFor(ctx);
      try {
        await client.post(`${msgUrl(ref.id)}/send`, undefined, { deadlineMs: SEND_DEADLINE_MS, now });
        return {};
      } catch (e) {
        if (isItemGone(e)) throw new Error('not sent: the Outlook draft no longer exists');
        throw sendError(e);
      }
    },
  };
}
