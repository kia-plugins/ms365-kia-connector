import type { ExternalRef } from '@kiagent/connector-sdk';
import { CAL_DOC_TYPE, calendarExternalId, isLiveEvent } from './document';
import { listCalendars, listView, type CalClient } from './graph';
import { DAY, HORIZON_DAYS, SINCE_DAYS, WINDOW_BACK_DAYS, type CalCursor } from './cursor';
import type { CalendarItem, CalMeta, GraphCalendar } from './types';

const iso = (ms: number) => new Date(ms).toISOString();
const utcStart = (dt?: string) => (dt ? new Date(`${dt.replace(/(\.\d{3})\d*$/, '$1')}Z`).toISOString() : '');

/** Explicit `config.calendarRoots` wins (an empty list = calendar off);
 *  absent = the default calendar plus every calendar the user owns. */
export function selectedCalendars(config: Record<string, unknown>, all: GraphCalendar[]): CalMeta[] {
  // Graph's hexColor may be 'auto' or '': only #rrggbb is a colour; null
  // lets the page pick one from its palette.
  const color = (h?: string) => (h && /^#[0-9a-f]{6}$/i.test(h) ? h : null);
  const meta = (c: GraphCalendar): CalMeta => ({ id: c.id, name: c.name, color: color(c.hexColor) });
  const roots = config.calendarRoots;
  if (Array.isArray(roots)) {
    const ids = new Set(roots.map((r) => (r as { id: string }).id));
    return all.filter((c) => ids.has(c.id)).map(meta);
  }
  return all.filter((c) => c.isDefaultCalendar || c.canEdit).map(meta);
}

export function calendarRootsConfigured(config: Record<string, unknown>): boolean {
  return Array.isArray(config.calendarRoots);
}

export async function pullCalendars(
  c: CalClient, config: Record<string, unknown>, prior: CalCursor | undefined, now: number,
): Promise<{ items: CalendarItem[]; deletions: ExternalRef[]; cursor: CalCursor }> {
  const since = prior?.since ?? iso(now - SINCE_DAYS * DAY);
  const windowStart = iso(Math.max(Date.parse(since), now - WINDOW_BACK_DAYS * DAY));
  const to = iso(now + HORIZON_DAYS * DAY);
  if (Array.isArray(config.calendarRoots) && config.calendarRoots.length === 0) {
    return { items: [], deletions: [], cursor: { since, cals: {} } };
  }
  const cals = selectedCalendars(config, await listCalendars(c));
  const items: CalendarItem[] = [];
  const deletions: ExternalRef[] = [];
  const next: CalCursor = { since, cals: {} };
  for (const cal of cals) {
    const known = prior?.cals[cal.id];
    const from = known ? windowStart : since; // a newly selected calendar gets its full history once
    const events = (await listView(c, cal.id, from, to)).filter(isLiveEvent);
    const seen: Record<string, string> = {};
    for (const e of events) {
      items.push({ calendar: cal, calendarEvent: e });
      const s = utcStart(e.start?.dateTime);
      if (s >= windowStart) seen[e.id] = s;
    }
    for (const [id, start] of Object.entries(known ?? {})) {
      if (start >= windowStart && !(id in seen)) {
        deletions.push({ externalId: calendarExternalId(cal.id, id), type: CAL_DOC_TYPE });
      }
    }
    next.cals[cal.id] = seen;
  }
  return { items, deletions, cursor: next };
}
