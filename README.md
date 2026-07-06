# Microsoft 365 connector for KIAgent

Indexes your Outlook mail (Microsoft 365 / Outlook.com) into your local
KIAgent digital memory via the Microsoft Graph API: each conversation becomes
one searchable email thread document, kept in sync automatically.

## Install

Install **Microsoft 365** from the KIAgent marketplace (Settings →
Extensions → Marketplace → Microsoft 365 → Install). KIAgent will ask for the
one grant this connector needs before it activates:

- `net` — to talk to `graph.microsoft.com`.

## Connect your account

1. Add a Microsoft 365 account. A Microsoft sign-in window opens — the OAuth
   flow (and the app registration behind it) is owned entirely by the
   platform. The connector requests only `Mail.Read` and `User.Read`, and
   never sees your Microsoft password; tokens live in KIAgent's encrypted
   vault and are refreshed by the platform. This extension ships **no**
   Microsoft client credentials of its own.
2. The account shows up under your Microsoft 365 mail address (or, if none
   is published, your sign-in name) and backfills your Inbox and Sent Items,
   then checks for changes every 15 minutes.

You can connect multiple Microsoft 365 accounts side by side — both personal
(Outlook.com/Hotmail) and work-or-school tenants.

## What gets indexed

- **Every conversation** in your **Inbox** and **Sent Items** — each becomes
  one `email.thread` document (the same document type the built-in Gmail
  source uses, so Outlook mail and Gmail mail render identically in the
  app), with all of the conversation's messages stacked into one markdown
  body, newest activity last.
- **Junk Email and Deleted Items are always excluded** from every sync —
  matching the legacy connector's behavior exactly.
- Messages upstream-deleted from your mailbox are archived out of the local
  index on the next sync.

## What does NOT get indexed (and why)

- **Attachments.** The legacy (v1) connector's Graph queries never actually
  requested the `attachments` field on a conversation's messages — its
  attachment-ingestion code path existed but was unreachable in production.
  This port preserves that OBSERVED behavior (no attachment sub-documents)
  rather than the apparently-intended-but-broken design; see the porting
  report for the full writeup. A future version could add real attachment
  support (Graph does expose attachment metadata and bytes cheaply) but that
  is new functionality, out of scope for a faithful v1→v2 port.
- **HTML-only message bodies** render with an empty body in the thread
  markdown. The legacy connector ran HTML through a local converter to
  markdown at ingest time; this pure `toDocument` mapping has no converter
  to call (the same simplification already made in the built-in Gmail v2
  port).
- Threads that legacy would have filtered as "automated" (auto-responders,
  mailing lists, bounce/DSN messages, `no-reply@…` senders, etc.) are **no
  longer filtered** — every conversation is indexed, matching the built-in
  Gmail source's v2 behavior.

## Privacy

- Read-only Mail scope; nothing is ever written to your mailbox.
- All content stays on your machine.
- This extension ships no Microsoft OAuth client credentials and stores no
  tokens itself — the platform's Microsoft OAuth provider owns the whole
  flow, including token refresh.

## Limitations

- Only the **Inbox** and **Sent Items** folders are synced — matching
  Microsoft Graph's own constraint that `/me/messages/delta` only supports
  change tracking scoped to a specific well-known folder for personal (MSA)
  accounts.
- No attachment ingestion (see above).
