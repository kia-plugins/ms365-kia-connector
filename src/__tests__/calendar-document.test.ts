import { calendarToDocument, isLiveEvent } from '../calendar/document';
import { descriptionFromMarkdown } from '../testing/description-from-markdown';

const cal = { id: 'CAL1', name: 'Calendar', color: '#0078d4' };
const base = {
  id: 'EV1', iCalUId: 'UID1', subject: 'Design review', type: 'singleInstance',
  start: { dateTime: '2026-10-01T09:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-10-01T10:00:00.0000000', timeZone: 'UTC' },
  isAllDay: false, isCancelled: false, showAs: 'busy',
  organizer: { emailAddress: { address: 'boss@x.com', name: 'Boss' } },
  attendees: [{ emailAddress: { address: 'me@x.com', name: 'Me' }, status: { response: 'accepted' }, type: 'required' }],
  responseStatus: { response: 'accepted' },
  onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/1' },
  location: { displayName: 'Room 1' },
  body: { contentType: 'text', content: 'Agenda:\n1. Q4' },
  originalStartTimeZone: 'W. Europe Standard Time',
};

test('a timed event maps to the shared metadata and markdown layout', () => {
  const d = calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: base as never });
  expect(d.type).toBe('calendar.event');
  expect(d.externalId).toBe('CAL1:EV1');
  expect(d.createdAt).toBe('2026-10-01T09:00:00.000Z');
  expect(d.scopeRootId).toBe('CAL1');
  expect(d.metadata).toMatchObject({
    calendarId: 'CAL1', calendarName: 'Calendar', calendarColor: '#0078d4',
    eventId: 'EV1', iCalUID: 'UID1', occurrenceKey: 'UID1',
    start: '2026-10-01T09:00:00.000Z', end: '2026-10-01T10:00:00.000Z', allDay: false,
    organizer: 'boss@x.com', selfResponse: 'accepted',
    attendees: [{ email: 'me@x.com', name: 'Me', response: 'accepted' }],
    participants: ['me@x.com', 'boss@x.com'],
    conferenceUrl: 'https://teams.microsoft.com/l/meetup-join/1', location: 'Room 1',
    transparency: 'opaque',
  });
  expect(d.url).toBe('https://outlook.office.com/calendar/item/EV1');
  expect(d.markdown!.split('\n')[0]).toBe('# Design review');
  expect(descriptionFromMarkdown(d.markdown!)).toBe('Agenda:\n1. Q4');
});

test('an all-day event carries local dates, end exclusive', () => {
  const d = calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: {
    ...base, isAllDay: true,
    start: { dateTime: '2026-10-01T00:00:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-10-03T00:00:00.0000000', timeZone: 'UTC' },
  } as never });
  expect(d.metadata).toMatchObject({ allDay: true, startDate: '2026-10-01', endDate: '2026-10-03' });
});

test('cancelled events are not live', () => {
  expect(isLiveEvent({ ...base, isCancelled: true } as never)).toBe(false);
  expect(isLiveEvent(base as never)).toBe(true);
});

test('an all-day event in a non-UTC calendar keeps its local dates (UTC-preferred times sit at 22:00 the day before)', () => {
  const d = calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: {
    ...base, isAllDay: true,
    start: { dateTime: '2026-09-30T22:00:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-10-02T22:00:00.0000000', timeZone: 'UTC' },
  } as never });
  expect(d.metadata).toMatchObject({ allDay: true, startDate: '2026-10-01', endDate: '2026-10-03' });
});

test('the description is the text body Graph returns (outlook.body-content-type="text")', () => {
  const d = calendarToDocument({ calendar: cal, tenantKind: 'work', calendarEvent: {
    ...base, body: { contentType: 'text', content: '  Agenda:\n1. Q4\n' },
  } as never });
  expect(descriptionFromMarkdown(d.markdown!)).toBe('Agenda:\n1. Q4');
});
