/**
 * Re-apply the standard spreadsheet formatting — bold header, frozen header
 * row, autoFilter, sensible column widths, wrapped text, clickable Job Link
 * cells — to any workbook whose first sheet has a header row.
 *
 * Ported from the former scripts/format_discovery.py (openpyxl). The data
 * writes themselves go through SheetJS (workbook.ts) for its date-serial /
 * number-format handling; SheetJS's cellStyles option is avoided there (see
 * the note in workbook.ts) because repeated read+write with rich styles
 * enabled corrupts this build's workbook theme part. exceljs is used only for
 * this cosmetic pass — open the file SheetJS already wrote, apply styling,
 * save — the same two-step shape the Python version had, just without Python.
 *
 * Two things here exist purely to keep Excel from raising its "We found a
 * problem with some content" repair prompt, and both are regressions of the
 * SheetJS-then-exceljs handoff rather than of either library alone:
 *
 * 1. `ensureValidSheetFormat` — SheetJS emits no <sheetFormatPr> element at
 *    all, so exceljs reads `worksheet.properties` back as `{}`. On write,
 *    exceljs then sees no defaultRowHeight, decides the height must be
 *    "custom", and emits `<sheetFormatPr customHeight="1"/>` — with the
 *    `defaultRowHeight` attribute the OOXML schema marks as REQUIRED missing.
 *    Excel treats that as an XML load error on the whole worksheet part
 *    ("Load error. Line 2, column 0" in its recovery log), repairs it by
 *    re-parsing leniently, and in the process drops every worksheet-level
 *    view/layout element (frozen pane, column widths, autoFilter) while the
 *    cell-level styling survives because it lives in the separate styles part.
 *    Giving exceljs Excel's own default (15pt) before writing is the fix.
 *
 * 2. `_xlnm._FilterDatabase` is stripped rather than maintained. It's the
 *    hidden, sheet-scoped defined name Excel itself creates alongside an
 *    autoFilter; exceljs's DefinedNames abstraction has no notion of a name's
 *    `localSheetId` scope, so it can't round-trip that name correctly. The
 *    <autoFilter> element in the sheet is sufficient on its own — Excel
 *    recreates the name the next time the user saves — so the simplest
 *    correct behaviour is to never carry a possibly-stale, possibly-misscoped
 *    copy forward.
 */

import ExcelJS from "exceljs";

const HEADER_FILL = "FF305496";
const HEADER_FONT_COLOR = "FFFFFFFF";
const LINK_COLOR = "FF0563C1";
const MIN_WIDTH = 12;
const MAX_WIDTH = 45;
const MAX_SAMPLE_LEN = 60;
/** Excel's own default row height, in points — what it writes for a fresh sheet. */
const DEFAULT_ROW_HEIGHT = 15;

function cellDisplayText(value: ExcelJS.CellValue): string {
  if (value == null) return "";
  if (typeof value === "object") {
    // Hyperlink ({text, hyperlink}), rich text ({richText: [...]}), or a Date.
    if (value instanceof Date) return value.toISOString();
    const v = value as any;
    if (typeof v.text === "string") return v.text;
    if (Array.isArray(v.richText)) return v.richText.map((r: any) => r.text ?? "").join("");
    return String(value);
  }
  return String(value);
}

/**
 * Make sure the sheet's <sheetFormatPr> will be written with the
 * schema-required `defaultRowHeight` attribute (see the header comment, #1).
 * Leaves any valid existing value alone.
 */
function ensureValidSheetFormat(ws: ExcelJS.Worksheet): void {
  const height = ws.properties?.defaultRowHeight;
  if (!(typeof height === "number" && Number.isFinite(height) && height > 0)) {
    ws.properties = { ...ws.properties, defaultRowHeight: DEFAULT_ROW_HEIGHT };
  }
}

export async function formatWorkbookFile(
  filePath: string
): Promise<{ ok: boolean; detail?: string }> {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.worksheets[0];

    // See header comment, #2.
    wb.definedNames.model = wb.definedNames.model.filter(
      (dn: { name: string }) => dn.name !== "_xlnm._FilterDatabase"
    );

    if (!ws) {
      await wb.xlsx.writeFile(filePath);
      return { ok: true };
    }

    ensureValidSheetFormat(ws);

    if (ws.rowCount < 1 || ws.columnCount < 1) {
      await wb.xlsx.writeFile(filePath);
      return { ok: true };
    }

    const colCount = ws.columnCount;
    const headerRow = ws.getRow(1);
    const headers: string[] = [];
    let linkCol = -1;
    for (let c = 1; c <= colCount; c++) {
      const cell = headerRow.getCell(c);
      const text = cellDisplayText(cell.value);
      headers[c - 1] = text;
      if (text.trim().toLowerCase() === "job link") linkCol = c;

      cell.font = { bold: true, color: { argb: HEADER_FONT_COLOR } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
      cell.alignment = { vertical: "middle", wrapText: true };
    }

    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: ws.rowCount, column: colCount } };

    const longest = new Array(colCount).fill(0);
    for (let c = 1; c <= colCount; c++) {
      longest[c - 1] = Math.min(headers[c - 1]?.length ?? 0, MAX_SAMPLE_LEN);
    }

    for (let r = 2; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= colCount; c++) {
        const cell = row.getCell(c);
        cell.alignment = { vertical: "top", wrapText: true };
        const text = cellDisplayText(cell.value);
        if (text) longest[c - 1] = Math.max(longest[c - 1], Math.min(text.length, MAX_SAMPLE_LEN));

        if (c === linkCol && typeof cell.value === "string") {
          const url = cell.value.trim();
          if (/^https?:\/\//i.test(url)) {
            cell.value = { text: url, hyperlink: url };
            cell.font = { color: { argb: LINK_COLOR }, underline: true };
          }
        }
      }
    }

    for (let c = 1; c <= colCount; c++) {
      ws.getColumn(c).width = Math.max(MIN_WIDTH, Math.min(longest[c - 1] + 2, MAX_WIDTH));
    }

    await wb.xlsx.writeFile(filePath);
    return { ok: true };
  } catch (err: any) {
    return { ok: false, detail: err?.message || String(err) };
  }
}
