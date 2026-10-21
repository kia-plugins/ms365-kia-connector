/**
 * Microsoft 365 CalendarWriter — adds events through Microsoft Graph.
 * Reachable only from kiagent-core's confirmation-gated booking pipeline.
 *
 *  - listCalendars: the account's calendars (Graph `canEdit` is the USER's
 *    right) with the mailbox time zone mapped to IANA.
 *  - createEvent: `POST /me/calendars/{id}/events`, idempotent through
 *    `transactionId` (derived from the intent's key), stamped with a named
 *    extended property so the indexed event keeps Kia's category.
 *
 * Failure wording is calendar-specific on purpose: the mail sender's
 * "grant send permission" copy would mislead here.
 */
import { createHash } from 'node:crypto';
import type {
  CalendarWriter,
  CreatedEvent,
  EventCategory,
  EventIntent,
  HostFor,
  SenderContext,
  WritableCalendar,
} from '@kiagent/connector-sdk';
import { GRAPH_BASE } from '../graph-api';
import { GraphApiError, GraphClient, type GraphClientDeps } from '../graph-client';
import { listCalendars } from './graph';
import { KIA_STAMP_ID } from './stamp';

/** Inside the host's call timeout; nothing may still be in flight after it. */
export const CREATE_DEADLINE_MS = 40_000;

const CATEGORY_NAME: Record<EventCategory, string> = { work: 'Work', personal: 'Personal', kids: 'Kids' };
/** Outlook's preset category colours. */
const CATEGORY_COLOR: Record<EventCategory, string> = { work: 'preset7', personal: 'preset4', kids: 'preset1' };

const WINDOWS_TO_IANA: Record<string, string> = {
  'W. Europe Standard Time': 'Europe/Berlin',
  'Central Europe Standard Time': 'Europe/Budapest',
  'Romance Standard Time': 'Europe/Paris',
  'Central European Standard Time': 'Europe/Warsaw',
  'GMT Standard Time': 'Europe/London',
  'FLE Standard Time': 'Europe/Kyiv',
  'E. Europe Standard Time': 'Europe/Chisinau',
  'Russian Standard Time': 'Europe/Moscow',
  'Turkey Standard Time': 'Europe/Istanbul',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Mountain Standard Time': 'America/Denver',
  'Central Standard Time': 'America/Chicago',
  'Eastern Standard Time': 'America/New_York',
  'India Standard Time': 'Asia/Kolkata',
  'China Standard Time': 'Asia/Shanghai',
  'Tokyo Standard Time': 'Asia/Tokyo',
  'AUS Eastern Standard Time': 'Australia/Sydney',
  UTC: 'UTC',
};

/** Windows name → IANA; an IANA-looking name passes through; else null. */
export function toIana(zone: string | undefined | null): string | null {
  if (!zone) return null;
  if (WINDOWS_TO_IANA[zone]) return WINDOWS_TO_IANA[zone];
  return /^[A-Za-z_]+(\/[A-Za-z0-9_+-]+)+$/.test(zone) ? zone : null;
}

/** A GUID (UUIDv5-shaped) derived from sha256(key): the same key always
 *  gives the same transactionId, so Graph drops a replayed create. */
export function transactionIdFor(key: string): string {
  const h = createHash('sha256').update(key).digest();
  h[6] = (h[6] & 0x0f) | 0x50;
  h[8] = (h[8] & 0x3f) | 0x80;
  const x = h.subarray(0, 16).toString('hex');
  return `${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20, 32)}`;
}

/** Wall-clock `YYYY-MM-DDTHH:mm:ss` of an instant in an IANA zone. */
function wallTime(iso: string, timeZone: string): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(new Date(iso));
  const p = (t: string) => parts.find((x) => x.type === t)!.value;
  return `${p('year')}-${p('month')}-${p('day')}T${p('hour')}:${p('minute')}:${p('second')}`;
}

const fold = (s: string | undefined | null) => (s ?? '').trim().toLowerCase();

export function createMs365CalendarWriter(
  host: HostFor<'net'>,
  clock?: Pick<GraphClientDeps, 'sleep' | 'random'>,
): CalendarWriter {
  const zoneByAccount = new Map<string, string | null>();
  const categoriesDone = new Set<string>();

  const clientFor = (ctx: SenderContext | undefined): GraphClient => {
    // Refreshed host-side per call — this process has no vault.
    const token = ctx?.credentials?.accessToken;
    if (!token) throw new Error('no Microsoft 365 credentials — reconnect the account in Settings');
    return new GraphClient({ fetch: host.net.fetch, getToken: async () => token, ...clock });
  };

  const myAddress = async (c: GraphClient): Promise<string | null> => {
    try {
      const me = await c.request<{ mail?: string | null; userPrincipalName?: string }>(
        `${GRAPH_BASE}/me?$select=mail,userPrincipalName`,
      );
      return me.mail || me.userPrincipalName || null;
    } catch {
      return null;
    }
  };

  const ensureCategories = async (c: GraphClient, accountKey: string): Promise<void> => {
    if (categoriesDone.has(accountKey)) return;
    categoriesDone.add(accountKey);
    try {
      const have = await c.request<{ value?: Array<{ displayName?: string }> }>(
        `${GRAPH_BASE}/me/outlook/masterCategories`,
      );
      const names = new Set((have.value ?? []).map((v) => fold(v.displayName)));
      for (const cat of ['work', 'personal', 'kids'] as const) {
        if (names.has(fold(CATEGORY_NAME[cat]))) continue;
        await c.write('POST', `${GRAPH_BASE}/me/outlook/masterCategories`, {
          body: { displayName: CATEGORY_NAME[cat], color: CATEGORY_COLOR[cat] },
          deadline: Date.now() + CREATE_DEADLINE_MS,
          idempotent: false,
        });
      }
    } catch (e) {
      // Never blocks the booking: the event is created either way.
      console.warn(`ms365: could not ensure Work/Personal/Kids categories: ${(e as Error).message}`);
    }
  };

  const failure = async (e: unknown, c: GraphClient): Promise<Error> => {
    if (!(e instanceof GraphApiError)) return e instanceof Error ? e : new Error(String(e));
    if (e.status === 403) {
      const who = await myAddress(c);
      return new Error(`reconnect Microsoft 365${who ? ` (${who})` : ''} in Settings to allow adding events`);
    }
    if (e.status === 429) return new Error('rate-limited: Outlook is busy — try again in a minute');
    if (e.status >= 400 && e.status < 500) {
      let msg: string | undefined;
      try {
        msg = (JSON.parse(e.message.slice(e.message.indexOf('{'))) as { error?: { message?: string } }).error?.message;
      } catch {
        /* snippet may be cut — fall back to the code */
      }
      return new Error(`not created: ${msg || e.code || `Outlook answered ${e.status}`}`);
    }
    return e;
  };

  return {
    async listCalendars(ctx: SenderContext): Promise<WritableCalendar[]> {
      const c = clientFor(ctx);
      const me = fold(await myAddress(c));
      let zone: string | null;
      if (zoneByAccount.has(me)) zone = zoneByAccount.get(me)!;
      else {
        try {
          const ms = await c.request<{ timeZone?: string }>(`${GRAPH_BASE}/me/mailboxSettings?$select=timeZone`);
          zone = toIana(ms.timeZone);
          zoneByAccount.set(me, zone);
        } catch {
          zone = null; // not cached: try again next time
        }
      }
      const all = await listCalendars(c);
      return all.map((cal) => ({
        id: cal.id,
        name: cal.name,
        color: cal.hexColor && /^#[0-9a-f]{6}$/i.test(cal.hexColor) ? cal.hexColor : null,
        primary: cal.isDefaultCalendar === true,
        canWrite: cal.canEdit === true,
        shared: !!cal.owner?.address && fold(cal.owner.address) !== me,
        timeZone: zone,
      }));
    },

    async createEvent(intent: EventIntent, ctx: SenderContext): Promise<CreatedEvent> {
      const c = clientFor(ctx);
      await ensureCategories(c, String(intent.accountId));

      let teams = false;
      if (intent.conference === 'auto') {
        try {
          const cal = await c.request<{ allowedOnlineMeetingProviders?: string[] }>(
            `${GRAPH_BASE}/me/calendars/${encodeURIComponent(intent.calendarId)}?$select=allowedOnlineMeetingProviders`,
          );
          teams = cal.allowedOnlineMeetingProviders?.includes('teamsForBusiness') === true;
        } catch {
          teams = false;
        }
      }

      const when = intent.allDay
        ? {
            start: { dateTime: `${intent.start}T00:00:00`, timeZone: intent.timeZone },
            end: { dateTime: `${intent.end}T00:00:00`, timeZone: intent.timeZone },
          }
        : {
            start: { dateTime: wallTime(intent.start, intent.timeZone), timeZone: intent.timeZone },
            end: { dateTime: wallTime(intent.end, intent.timeZone), timeZone: intent.timeZone },
          };
      const payload = {
        subject: intent.title,
        body: { contentType: 'text', content: intent.description ?? '' },
        ...when,
        isAllDay: intent.allDay,
        ...(intent.location ? { location: { displayName: intent.location } } : {}),
        attendees: intent.attendees.map((a) => ({
          emailAddress: { address: a.email, ...(a.name ? { name: a.name } : {}) },
          type: 'required',
        })),
        categories: [CATEGORY_NAME[intent.category]],
        transactionId: transactionIdFor(intent.idempotencyKey),
        singleValueExtendedProperties: [{ id: KIA_STAMP_ID, value: `${intent.category}|${intent.kia.draftId}` }],
        ...(teams ? { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' } : {}),
      };

      let ev: { id: string; iCalUId?: string; webLink?: string; onlineMeeting?: { joinUrl?: string } | null };
      try {
        ev = await c.write('POST', `${GRAPH_BASE}/me/calendars/${encodeURIComponent(intent.calendarId)}/events`, {
          body: payload,
          deadline: Date.now() + CREATE_DEADLINE_MS,
          idempotent: true,
        });
      } catch (e) {
        throw await failure(e, c);
      }
      return {
        eventId: ev.id,
        iCalUID: ev.iCalUId ?? null,
        htmlLink: ev.webLink ?? null,
        conferenceUrl: ev.onlineMeeting?.joinUrl ?? null,
        conferencePending: false,
        note:
          intent.conference === 'auto' && !teams
            ? "no video link: this Outlook account can't create Teams meetings"
            : null,
      };
    },
  };
}
