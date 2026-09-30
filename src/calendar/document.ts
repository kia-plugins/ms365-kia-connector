import type { DocumentInput } from '@kiagent/connector-sdk';
import type { CalendarItem, GraphDateTime, GraphEvent } from './types';

export const CAL_DOC_TYPE = 'calendar.event';
export const calendarExternalId = (calId: string, eventId: string): string => `${calId}:${eventId}`;
export const isLiveEvent = (e: GraphEvent): boolean => e.isCancelled !== true;

/** Graph returns `2026-10-01T09:00:00.0000000` (no zone) in the requested
 *  zone, UTC here. */
const utc = (t?: GraphDateTime): string | null =>
  t?.dateTime ? new Date(`${t.dateTime.replace(/(\.\d{3})\d*$/, '$1')}Z`).toISOString() : null;
const datePart = (t?: GraphDateTime): string | undefined => t?.dateTime?.slice(0, 10);

const RESPONSE: Record<string, string> = {
  accepted: 'accepted', tentativelyAccepted: 'tentative', declined: 'declined',
  notResponded: 'needsAction', none: 'needsAction', organizer: 'accepted',
};
const response = (r?: string): string | null => (r ? RESPONSE[r] ?? r : null);

const text = (e: GraphEvent): string => {
  const c = e.body?.content ?? '';
  if (e.body?.contentType !== 'html') return c.trim();
  return c.replace(/<br\s*\/?>/gi, '\n').replace(/<\/p>/gi, '\n').replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
};

export function calendarToDocument(item: CalendarItem): DocumentInput {
  const { calendar: cal, calendarEvent: e } = item;
  const allDay = e.isAllDay === true;
  const start = utc(e.start);
  const end = utc(e.end);
  const attendees = (e.attendees ?? [])
    .filter((a) => a.type !== 'resource' && a.emailAddress?.address)
    .map((a) => ({ email: a.emailAddress!.address!, name: a.emailAddress?.name ?? null, response: response(a.status?.response) }));
  const organizer = e.organizer?.emailAddress?.address ?? null;
  const participants = [...new Set([...attendees.map((a) => a.email), ...(organizer ? [organizer] : [])])];
  const title = e.subject || '(no title)';
  const conferenceUrl = e.onlineMeeting?.joinUrl ?? null;
  const location = e.location?.displayName || null;
  const when = allDay ? `${datePart(e.start)} – ${datePart(e.end)} (all day)` : `${start} – ${end}`;
  const description = text(e);
  // Layout is the calendar.event contract (spec §2): title block, header
  // block, then the description — the assistant's brief reads it.
  const lines = [
    `# ${title}`, '',
    `**When:** ${when}`,
    ...(location ? [`**Where:** ${location}`] : []),
    `**Calendar:** ${cal.name}`,
    ...(organizer ? [`**Organizer:** ${organizer}`] : []),
    ...(attendees.length ? [`**Attendees:** ${attendees.map((a) => `${a.name ? `${a.name} ` : ''}<${a.email}> (${a.response ?? 'unknown'})`).join(', ')}`] : []),
    ...(conferenceUrl ? [`**Conference:** ${conferenceUrl}`] : []),
    ...(description ? ['', description] : []),
  ];
  return {
    externalId: calendarExternalId(cal.id, e.id),
    type: CAL_DOC_TYPE,
    title,
    markdown: lines.join('\n'),
    url: `https://outlook.office.com/calendar/item/${encodeURIComponent(e.id)}`,
    createdAt: start,
    scopeRootId: cal.id,
    metadata: {
      calendarId: cal.id, calendarName: cal.name, calendarColor: cal.color,
      eventId: e.id, iCalUID: e.iCalUId ?? null,
      // Graph's iCalUId is already per occurrence (spec §2).
      occurrenceKey: e.iCalUId ?? e.id,
      start, end, allDay,
      ...(allDay ? { startDate: datePart(e.start), endDate: datePart(e.end) } : {}),
      timeZone: e.originalStartTimeZone ?? 'UTC',
      status: e.isCancelled ? 'cancelled' : 'confirmed',
      organizer, selfResponse: response(e.responseStatus?.response),
      attendees, participants, conferenceUrl, location,
      eventType: 'default',
      transparency: e.showAs === 'free' ? 'transparent' : 'opaque',
    },
  };
}
