/**
 * Shared identity/dedupe helpers for Company+Position matching across the
 * discovery sheet, the tracker, and the board sweep. Punctuation must not
 * decide identity — an ATS posting "Director of Engineering - Booking Group"
 * and a tracker row "Director of Engineering, Booking Group" are the same
 * role — so everything non-alphanumeric collapses to a single space.
 */

/** Identity key: lowercase, punctuation/dashes flattened to spaces. */
export function norm(s: string | undefined | null): string {
  return (s ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function normUrl(u: string | undefined | null): string {
  return (u ?? "").trim().toLowerCase().replace(/[/?#]+$/, "");
}

/** Every http(s) link found in a "Job Link" cell (pipe-separated allowed). */
export function linksIn(cell: string | undefined): Set<string> {
  const out = new Set<string>();
  for (const piece of (cell ?? "").split("|")) {
    const token = piece.trim().split(" ")[0];
    if (token && token.toLowerCase().startsWith("http")) out.add(normUrl(token));
  }
  return out;
}

export function keyOf(company: string, position: string): string {
  return `${norm(company)} ${norm(position)}`;
}
