/**
 * Workbook read/write engine + shared fs utilities.
 *
 * Schema-agnostic sheet access: open a workbook fresh on every call, read/write
 * cells preserving number formats and styles, and back up before mutating.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import * as XLSX from "xlsx";
import { UserFacingError, wrapFsError } from "./errors.js";
import { serialToISO, isoToSerial } from "./dates.js";
import { SheetSpec, COLUMNS, isDateColumn, TRACKER } from "./config.js";

// SheetJS's ESM build does not auto-wire Node's fs the way the CJS build does,
// so readFile/writeFile would fail with "Cannot access file". Register it once.
XLSX.set_fs(fs);

/* ------------------------------------------------------------------ */
/* Shared fs utilities                                                */
/* ------------------------------------------------------------------ */

export function ensureDir(dir: string): void {
  try {
    fs.mkdirSync(dir, { recursive: true });
  } catch (err) {
    throw wrapFsError(err, "create the folder for", dir);
  }
}

export function humanSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}

/* ------------------------------------------------------------------ */
/* Workbook access                                                    */
/* ------------------------------------------------------------------ */

export interface SheetHandle {
  /** which workbook this handle is for (path, schema, date columns, …). */
  spec: SheetSpec;
  wb: XLSX.WorkBook;
  ws: XLSX.WorkSheet;
  sheetName: string;
  /**
   * All columns present in the header row, in sheet order. Base columns use
   * their canonical spelling; extra (custom) columns use their header text.
   */
  order: string[];
  /** column display-name -> 0-based column index */
  colIndex: Map<string, number>;
  range: XLSX.Range;
  /** 0-based row index of the first data row (header + 1). */
  firstDataRow: number;
  /** 0-based row index of the last data row (may be < firstDataRow if empty). */
  lastDataRow: number;
}

export interface JobRecord {
  /** 0-based sheet row index of this record. */
  row: number;
  /** column display-name -> cell text, for every column in the sheet. */
  values: Record<string, string>;
  /** raw serials for date columns, when present (for computation). */
  serials: Record<string, number>;
}

/** Case-insensitive column-index lookup. Returns undefined if absent. */
export function indexOf(h: SheetHandle, name: string): number | undefined {
  const want = name.trim().toLowerCase();
  for (const [key, idx] of h.colIndex) {
    if (key.toLowerCase() === want) return idx;
  }
  return undefined;
}

/** Does the sheet have a column with this (case-insensitive) name? */
export function hasColumn(h: SheetHandle, name: string): boolean {
  return indexOf(h, name) !== undefined;
}

/** Read a workbook fresh, with clear errors for the common failure modes. */
export function openWorkbook(spec: SheetSpec = TRACKER): SheetHandle {
  const filePath = spec.filePath;
  if (!fs.existsSync(filePath)) {
    const envVar =
      spec.label === "discovery" ? "JOB_DISCOVERY_FILE" : "JOB_TRACKER_FILE";
    throw new UserFacingError(
      `The ${spec.label} spreadsheet was not found at:\n  ${filePath}\n` +
        `Set the ${envVar} environment variable to point at it.`
    );
  }

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.readFile(filePath, { cellStyles: true, cellNF: true });
  } catch (err) {
    throw wrapFsError(err, "read", filePath);
  }

  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws || !ws["!ref"]) {
    throw new UserFacingError(
      `The first sheet ("${sheetName}") in ${path.basename(filePath)} is empty ` +
        `or has no data range.`
    );
  }

  const range = XLSX.utils.decode_range(ws["!ref"]);

  // Read every header in sheet order. Base columns are normalised to their
  // canonical spelling; any extra columns keep their own header text. This is
  // what makes the schema dynamic — extras become first-class alongside the base.
  const colIndex = new Map<string, number>();
  const order: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
    const header = cell && cell.v != null ? String(cell.v).trim() : "";
    if (!header) continue; // skip blank header cells
    const canonical = spec.baseColumns.find(
      (col) => col.toLowerCase() === header.toLowerCase()
    );
    const key = canonical ?? header;
    // First occurrence wins; ignore accidental duplicate headers.
    if (colIndex.has(key)) continue;
    colIndex.set(key, c);
    order.push(key);
  }

  const missing = spec.baseColumns.filter((c) => !colIndex.has(c));
  if (missing.length) {
    throw new UserFacingError(
      `The ${spec.label} spreadsheet is missing required column(s): ` +
        `${missing.join(", ")}.\nExpected header row: ${spec.baseColumns.join(" | ")}`
    );
  }

  return {
    spec,
    wb,
    ws,
    sheetName,
    order,
    colIndex,
    range,
    firstDataRow: range.s.r + 1,
    lastDataRow: range.e.r,
  };
}

/** Read a single cell's display string + raw serial (for date columns). */
export function readCell(
  h: SheetHandle,
  row: number,
  col: string
): { text: string; serial?: number } {
  const c = indexOf(h, col);
  if (c === undefined) return { text: "" };
  const cell = h.ws[XLSX.utils.encode_cell({ r: row, c })];
  if (!cell || cell.v == null || cell.v === "") return { text: "" };

  if (isDateColumn(h.spec, col)) {
    if (cell.t === "n" && typeof cell.v === "number") {
      return { text: serialToISO(cell.v), serial: cell.v };
    }
    // Fall back to whatever text is there (e.g. a hand-typed date string).
    return { text: String(cell.w ?? cell.v).trim() };
  }
  return { text: String(cell.w ?? cell.v).trim() };
}

/** Is a data row entirely blank across every column? (Skips trailing empties.) */
export function isRowBlank(h: SheetHandle, row: number): boolean {
  return h.order.every((col) => readCell(h, row, col).text === "");
}

/** Read every non-blank data row into structured records (all columns). */
export function readAllRecords(h: SheetHandle): JobRecord[] {
  const out: JobRecord[] = [];
  for (let r = h.firstDataRow; r <= h.lastDataRow; r++) {
    if (isRowBlank(h, r)) continue;
    const values: Record<string, string> = {};
    const serials: Record<string, number> = {};
    for (const col of h.order) {
      const { text, serial } = readCell(h, r, col);
      values[col] = text;
      if (serial !== undefined) serials[col] = serial;
    }
    out.push({ row: r, values, serials });
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Writing                                                            */
/* ------------------------------------------------------------------ */

/** How many timestamped backups to keep per file. */
const BACKUP_KEEP = 10;

/** Local timestamp "YYYYMMDD-HHmmss-SSS" for backup filenames. */
export function backupStamp(): string {
  const d = new Date();
  const p = (n: number, w = 2) => String(n).padStart(w, "0");
  return (
    `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-` +
    `${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}-${p(
      d.getMilliseconds(),
      3
    )}`
  );
}

/**
 * Copy `filePath` into a sibling ".backups" folder with a timestamped name,
 * keeping only the most recent BACKUP_KEEP backups for that file. This replaces
 * the old single-slot "<file>.bak" scheme so repeated writes (manual or the
 * daily task) can't clobber the last good copy. Returns the backup path.
 */
export function backupWithRotation(filePath: string): string {
  const dir = path.join(path.dirname(filePath), ".backups");
  ensureDir(dir);
  const base = path.basename(filePath);
  const dest = path.join(dir, `${base}.${backupStamp()}.bak`);
  try {
    fs.copyFileSync(filePath, dest);
  } catch (err) {
    throw wrapFsError(err, "back up", filePath);
  }
  // Prune older backups for this file, keeping the newest BACKUP_KEEP.
  try {
    const prefix = base + ".";
    const olds = fs
      .readdirSync(dir)
      .filter((f) => f.startsWith(prefix) && f.endsWith(".bak"))
      .map((f) => ({ f, t: fs.statSync(path.join(dir, f)).mtimeMs }))
      .sort((a, b) => b.t - a.t)
      .slice(BACKUP_KEEP);
    for (const o of olds) {
      try {
        fs.unlinkSync(path.join(dir, o.f));
      } catch {
        /* ignore individual prune failures */
      }
    }
  } catch {
    /* ignore prune errors entirely — backup already succeeded */
  }
  return dest;
}

export function backupFile(h: SheetHandle): string {
  return backupWithRotation(h.spec.filePath);
}

export function saveWorkbook(h: SheetHandle): void {
  try {
    XLSX.writeFile(h.wb, h.spec.filePath, { cellStyles: true });
  } catch (err) {
    throw wrapFsError(err, "write", h.spec.filePath);
  }
}

/**
 * Set a cell, copying number format / style from a template cell in the same
 * column (the last existing data row) so new rows look like the ones above —
 * the way a human copy-pasting a row would leave them.
 */
export function writeCell(
  h: SheetHandle,
  row: number,
  col: string,
  value: string,
  templateRow: number
): void {
  const c = indexOf(h, col);
  if (c === undefined) {
    throw new UserFacingError(`Internal error: no column "${col}" to write to.`);
  }
  const addr = XLSX.utils.encode_cell({ r: row, c });
  const templateAddr = XLSX.utils.encode_cell({ r: templateRow, c });
  const template = h.ws[templateAddr] as XLSX.CellObject | undefined;

  if (value === "" || value == null) {
    // Leave optional blanks unset; they inherit column formatting anyway.
    delete h.ws[addr];
    return;
  }

  let cell: XLSX.CellObject;
  if (isDateColumn(h.spec, col)) {
    const serial = isoToSerial(value);
    cell = { t: "n", v: serial, z: (template && template.z) || "d-mmm-yy" };
  } else {
    cell = { t: "s", v: String(value) };
    if (template && template.z) cell.z = template.z;
  }
  // Best-effort style carry-over (cell fonts/fills). Harmless if unsupported.
  if (template && (template as any).s) (cell as any).s = (template as any).s;

  h.ws[addr] = cell;
}

/** Extend the sheet's !ref so it includes `row`. */
export function growRange(h: SheetHandle, row: number): void {
  const r = XLSX.utils.decode_range(h.ws["!ref"]!);
  if (row > r.e.r) r.e.r = row;
  h.ws["!ref"] = XLSX.utils.encode_range(r);
}

/**
 * Append a new data row from a list of {column, value} writes, copying number
 * formats / styles from the last existing row (like a human copy-pasting a row).
 * Column names are resolved case-insensitively; unknown columns are skipped by
 * writeCell's own guard, so callers should validate names first. Does not save.
 * Returns the 0-based index of the new row.
 */
export function appendRow(
  h: SheetHandle,
  writes: { col: string; value: string }[]
): number {
  const records = readAllRecords(h);
  const templateRow =
    records.length > 0 ? records[records.length - 1].row : h.firstDataRow - 1;
  const newRow =
    records.length > 0 ? records[records.length - 1].row + 1 : h.firstDataRow;
  for (const w of writes) {
    if (hasColumn(h, w.col)) writeCell(h, newRow, w.col, w.value, templateRow);
  }
  growRange(h, newRow);
  return newRow;
}

/**
 * Delete a whole data row and shift the rows below it up by one — the way
 * "Delete row" in Excel behaves (not merely clearing cells). Cell objects are
 * moved intact, so number formats / styles ride along with the rows above.
 */
export function deleteRow(h: SheetHandle, row: number): void {
  const r = XLSX.utils.decode_range(h.ws["!ref"]!);
  for (let cur = row; cur < r.e.r; cur++) {
    for (let c = r.s.c; c <= r.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r: cur + 1, c });
      const to = XLSX.utils.encode_cell({ r: cur, c });
      if (h.ws[from]) h.ws[to] = h.ws[from];
      else delete h.ws[to];
    }
  }
  // Clear the now-vacated last row.
  for (let c = r.s.c; c <= r.e.c; c++) {
    delete h.ws[XLSX.utils.encode_cell({ r: r.e.r, c })];
  }
  // Shrink the range by one row (never above the header row).
  r.e.r = Math.max(r.s.r, r.e.r - 1);
  h.ws["!ref"] = XLSX.utils.encode_range(r);
}

/* ------------------------------------------------------------------ */
/* Column-structure helpers (arbitrary headers, not just canonical)   */
/* ------------------------------------------------------------------ */

/** Lowercased canonical headers; these are structural and protected. */
export const CANONICAL_LC = new Set<string>(COLUMNS.map((c) => c.toLowerCase()));

export interface HeaderCell {
  /** trimmed header text ("" for a blank header cell) */
  name: string;
  /** 0-based column index */
  col: number;
}

/** Every header cell in the sheet, in column order (re-read from !ref). */
export function readHeaderRow(h: SheetHandle): HeaderCell[] {
  const rng = XLSX.utils.decode_range(h.ws["!ref"]!);
  const headerRow = rng.s.r;
  const out: HeaderCell[] = [];
  for (let c = rng.s.c; c <= rng.e.c; c++) {
    const cell = h.ws[XLSX.utils.encode_cell({ r: headerRow, c })];
    const name = cell && cell.v != null ? String(cell.v).trim() : "";
    out.push({ name, col: c });
  }
  return out;
}

/** Find a header by (case-insensitive) name, ignoring blank header cells. */
export function findHeader(h: SheetHandle, name: string): HeaderCell | undefined {
  const want = name.trim().toLowerCase();
  if (!want) return undefined;
  return readHeaderRow(h).find((x) => x.name.toLowerCase() === want);
}

/** Reject operations that would break the fixed schema the other tools rely on. */
export function assertNotCanonical(name: string, verb: string): void {
  if (CANONICAL_LC.has(name.trim().toLowerCase())) {
    throw new UserFacingError(
      `"${name}" is one of the built-in tracker columns and cannot be ${verb} — ` +
        `the other tools depend on it. You can only ${verb} custom columns you added. ` +
        `Built-in columns: ${COLUMNS.join(", ")}.`
    );
  }
}

/**
 * Delete a whole column and shift the columns to its right left by one — the
 * way "Delete column" in Excel behaves. Cell objects move intact so formats /
 * styles ride along. Keeps the "!cols" width array aligned.
 */
export function deleteColumn(h: SheetHandle, col: number): void {
  const r = XLSX.utils.decode_range(h.ws["!ref"]!);
  for (let cur = col; cur < r.e.c; cur++) {
    for (let row = r.s.r; row <= r.e.r; row++) {
      const from = XLSX.utils.encode_cell({ r: row, c: cur + 1 });
      const to = XLSX.utils.encode_cell({ r: row, c: cur });
      if (h.ws[from]) h.ws[to] = h.ws[from];
      else delete h.ws[to];
    }
  }
  // Clear the now-vacated last column.
  for (let row = r.s.r; row <= r.e.r; row++) {
    delete h.ws[XLSX.utils.encode_cell({ r: row, c: r.e.c })];
  }
  // Keep column widths aligned with the columns that remain.
  if (Array.isArray(h.ws["!cols"]) && h.ws["!cols"].length > col) {
    h.ws["!cols"].splice(col, 1);
  }
  // Shrink the range by one column (never past the first column).
  r.e.c = Math.max(r.s.c, r.e.c - 1);
  h.ws["!ref"] = XLSX.utils.encode_range(r);
}

/**
 * Validate an `extra_fields` map (custom columns → values) and resolve it to a
 * list of writes. Rejects unknown names (with a helpful list of the custom
 * columns that do exist) and base columns (which have their own parameters).
 */
export function resolveExtraFields(
  h: SheetHandle,
  extra: Record<string, string> | undefined
): { col: string; value: string }[] {
  const out: { col: string; value: string }[] = [];
  if (!extra) return out;
  for (const [name, value] of Object.entries(extra)) {
    if (CANONICAL_LC.has(name.trim().toLowerCase())) {
      throw new UserFacingError(
        `"${name}" is a built-in column — set it with its own parameter, ` +
          `not extra_fields.`
      );
    }
    if (!hasColumn(h, name)) {
      const custom = h.order.filter((c) => !CANONICAL_LC.has(c.toLowerCase()));
      throw new UserFacingError(
        `No column named "${name}". ` +
          (custom.length
            ? `Custom columns: ${custom.join(", ")}. `
            : "There are no custom columns yet. ") +
          `Create it first with add_column.`
      );
    }
    out.push({ col: name, value: String(value) });
  }
  return out;
}

/**
 * Resolve a generic `fields` map (any existing column → value) to writes.
 * Used by the discovery tools, where every column is set through `fields`.
 * Rejects unknown columns and any name in `reject` (columns that have their
 * own dedicated parameter, e.g. Company / Position).
 */
export function resolveFieldWrites(
  h: SheetHandle,
  fields: Record<string, string> | undefined,
  reject: ReadonlySet<string>
): { col: string; value: string }[] {
  const out: { col: string; value: string }[] = [];
  if (!fields) return out;
  for (const [name, value] of Object.entries(fields)) {
    if (reject.has(name.trim().toLowerCase())) {
      throw new UserFacingError(
        `Set "${name}" with its own parameter, not fields.`
      );
    }
    if (!hasColumn(h, name)) {
      throw new UserFacingError(
        `The ${h.spec.label} sheet has no column "${name}". ` +
          `Columns: ${h.order.join(", ")}.`
      );
    }
    out.push({ col: name, value: String(value) });
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Matching helpers                                                   */
/* ------------------------------------------------------------------ */

export function eq(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}
export function includesCI(haystack: string, needle: string): boolean {
  return haystack.toLowerCase().includes(needle.trim().toLowerCase());
}

export function findByCompanyPosition(
  records: JobRecord[],
  company: string,
  position?: string
): JobRecord[] {
  let matches = records.filter((r) => eq(r.values["Company"], company));
  if (position && position.trim()) {
    matches = matches.filter((r) => eq(r.values["Position"], position));
  }
  return matches;
}

/* ------------------------------------------------------------------ */
/* Presentation                                                       */
/* ------------------------------------------------------------------ */

/** A compact object for a record, dropping empty fields for readability.
 *  Includes every column in the sheet (base + custom), in sheet order. */
export function present(
  h: SheetHandle,
  rec: JobRecord,
  extra?: Record<string, unknown>
) {
  const obj: Record<string, unknown> = {};
  for (const col of h.order) {
    if (rec.values[col] !== undefined && rec.values[col] !== "") {
      obj[col] = rec.values[col];
    }
  }
  return { ...obj, ...(extra ?? {}) };
}
