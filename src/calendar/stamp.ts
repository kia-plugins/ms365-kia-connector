import type { GraphEvent } from './types';

/** Named single-value extended property Kia stamps on every event it
 *  creates: `<category>|<draftId>`. Read back on the calendarView so the
 *  indexed event keeps its Work/Personal/Kids category. */
export const KIA_STAMP_ID = 'String {f5d1d8b2-3c4e-4a7b-9d61-6b1c2a9e7f10} Name kiaStamp';

const CATEGORIES = ['work', 'personal', 'kids'] as const;
type Category = (typeof CATEGORIES)[number];
const asCategory = (s: string | undefined): Category | null => {
  const v = (s ?? '').trim().toLowerCase();
  return (CATEGORIES as readonly string[]).includes(v) ? (v as Category) : null;
};

/** The stamp's category, else the first of the event's Outlook categories
 *  named work|personal|kids (case-insensitive), else null. */
export function categoryOf(e: GraphEvent): Category | null {
  const stamp = e.singleValueExtendedProperties?.find((p) => p.id.toLowerCase() === KIA_STAMP_ID.toLowerCase());
  const fromStamp = asCategory(stamp?.value?.split('|')[0]);
  if (fromStamp) return fromStamp;
  for (const c of e.categories ?? []) {
    const hit = asCategory(c);
    if (hit) return hit;
  }
  return null;
}
