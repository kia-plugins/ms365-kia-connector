import { GRAPH_BASE } from '../graph-api';
import type { GraphCalendar, GraphEvent } from './types';

export interface CalClient {
  request<T>(url: string, opts?: { extraHeaders?: Record<string, string> }): Promise<T>;
}
const UTC = { prefer: 'outlook.timezone="UTC"' };

export async function listCalendars(c: CalClient): Promise<GraphCalendar[]> {
  const out: GraphCalendar[] = [];
  let url: string | undefined = `${GRAPH_BASE}/me/calendars?$select=id,name,hexColor,isDefaultCalendar,canEdit,owner&$top=100`;
  while (url) {
    const page: { value?: GraphCalendar[]; '@odata.nextLink'?: string } = await c.request(url);
    out.push(...(page.value ?? []));
    url = page['@odata.nextLink'];
  }
  return out;
}

export async function listView(c: CalClient, calId: string, from: string, to: string): Promise<GraphEvent[]> {
  const out: GraphEvent[] = [];
  let url: string | undefined =
    `${GRAPH_BASE}/me/calendars/${encodeURIComponent(calId)}/calendarView` +
    `?startDateTime=${encodeURIComponent(from)}&endDateTime=${encodeURIComponent(to)}&$top=250`;
  while (url) {
    const page: { value?: GraphEvent[]; '@odata.nextLink'?: string } = await c.request(url, { extraHeaders: UTC });
    out.push(...(page.value ?? []));
    url = page['@odata.nextLink'];
  }
  return out;
}
