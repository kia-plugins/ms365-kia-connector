/** Verbatim copy of alpha-cent extensions/assistant/src/triggers/calendar.ts
 *  `descriptionFromMarkdown`: the pre-meeting brief reads the event
 *  description this way, so the calendar.event markdown layout must keep
 *  working with it (spec 2026-09-30 §2). */
export function descriptionFromMarkdown(md: string | null): string | null {
  if (!md) return null;
  const blocks = md.split('\n\n');
  if (blocks.length < 3) return null;
  const text = blocks.slice(2).join('\n\n').trim();
  return text || null;
}
