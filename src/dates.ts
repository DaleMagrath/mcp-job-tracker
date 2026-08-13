/**
 * Date helpers (Excel serial <-> calendar date, TZ-free).
 */

import { UserFacingError } from "./errors.js";

export const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30); // 1899-12-30
export const MS_PER_DAY = 86_400_000;

/** Excel date serial -> "YYYY-MM-DD" (UTC calendar date, no time zone drift). */
export function serialToISO(serial: number): string {
  const ms = EXCEL_EPOCH_UTC + Math.round(serial) * MS_PER_DAY;
  const d = new Date(ms);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** "YYYY-MM-DD" (or a Date) -> Excel date serial. */
export function isoToSerial(iso: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso.trim());
  if (!m) {
    throw new UserFacingError(
      `Invalid date "${iso}". Please use ISO format YYYY-MM-DD (e.g. 2026-07-28).`
    );
  }
  const [, y, mo, d] = m;
  const ms = Date.UTC(Number(y), Number(mo) - 1, Number(d));
  return Math.round((ms - EXCEL_EPOCH_UTC) / MS_PER_DAY);
}

/** Today's local calendar date as an Excel serial and as ISO. */
export function todaySerial(): number {
  const now = new Date();
  const ms = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.round((ms - EXCEL_EPOCH_UTC) / MS_PER_DAY);
}
export function todayISO(): string {
  return serialToISO(todaySerial());
}

/** Whole days between an applied serial and today (today - applied). */
export function daysSince(serial: number): number {
  return todaySerial() - Math.round(serial);
}
