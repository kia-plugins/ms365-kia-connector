export interface GraphDateTime { dateTime: string; timeZone?: string }
export interface GraphEmail { address?: string; name?: string }
export interface GraphEvent {
  id: string;
  iCalUId?: string;
  subject?: string | null;
  type?: 'singleInstance' | 'occurrence' | 'exception' | 'seriesMaster';
  start?: GraphDateTime;
  end?: GraphDateTime;
  isAllDay?: boolean;
  isCancelled?: boolean;
  showAs?: string;
  organizer?: { emailAddress?: GraphEmail };
  attendees?: Array<{ emailAddress?: GraphEmail; status?: { response?: string }; type?: string }>;
  responseStatus?: { response?: string };
  onlineMeeting?: { joinUrl?: string } | null;
  location?: { displayName?: string } | null;
  body?: { contentType?: string; content?: string } | null;
  originalStartTimeZone?: string;
}
export interface CalMeta { id: string; name: string; color: string | null }
export interface CalendarItem { calendar: CalMeta; calendarEvent: GraphEvent }
export interface GraphCalendar { id: string; name: string; hexColor?: string; isDefaultCalendar?: boolean; canEdit?: boolean; owner?: GraphEmail }
