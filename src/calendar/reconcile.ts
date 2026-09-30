import type { ExternalRef } from '@kiagent/connector-sdk';
import { statusOf } from '../graph-client';
import { CAL_DOC_TYPE, calendarExternalId, isLiveEvent } from './document';
import { listCalendars, listView, type CalClient } from './graph';
import { DAY, HORIZON_DAYS, type CalCursor } from './cursor';
import { selectedCalendars } from './pull';

/** Every live event the pull may have stored: each selected calendar from
 *  the fixed first-pull start (spec §6). A 403 is "no events stored" ONLY
 *  when calendars never synced; otherwise any failure fails the pass. */
export async function calendarRefs(
  c: CalClient, config: Record<string, unknown>, cal: CalCursor | undefined, now: number,
): Promise<ExternalRef[]> {
  if (!cal) {
    try { await listCalendars(c); } catch (e) { if (statusOf(e) === 403) return []; throw e; }
    return []; // consent exists but calendars never synced: nothing stored yet
  }
  const to = new Date(now + HORIZON_DAYS * DAY).toISOString();
  const refs: ExternalRef[] = [];
  for (const m of selectedCalendars(config, await listCalendars(c))) {
    for (const e of (await listView(c, m.id, cal.since, to)).filter(isLiveEvent)) {
      refs.push({ externalId: calendarExternalId(m.id, e.id), type: CAL_DOC_TYPE });
    }
  }
  return refs;
}
