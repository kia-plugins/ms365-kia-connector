export interface CalCursor {
  /** First-pull start: fixed forever. Reconcile lists from here (spec §6). */
  since: string;
  /** calendar id → (event id → start ISO) for events inside the sliding window. */
  cals: Record<string, Record<string, string>>;
}
export const DAY = 86_400_000;
export const SINCE_DAYS = 365;
export const WINDOW_BACK_DAYS = 35;
export const HORIZON_DAYS = 400;
