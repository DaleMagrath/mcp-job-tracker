/**
 * Re-apply the standard spreadsheet formatting — bold header, frozen top row,
 * auto-filter, wrapped text, sensible column widths, clickable Job Link
 * cells — to any workbook whose first sheet has a header row.
 *
 * Ported from the former scripts/format_discovery.py (openpyxl). The data
 * writes themselves go through SheetJS (workbook.ts) for its date-serial /
 * number-format handling; SheetJS's cellStyles option is avoided there (see
 * the note in workbook.ts) because repeated read+write with rich styles
 * enabled corrupts this build's workbook theme part. exceljs is used only for
 * this cosmetic pass — open the file SheetJS already wrote, apply styling,
 * save — the same two-step shape the Python version had, just without Python.
 */

import ExcelJS from "exceljs";

const HEADER_FILL = "FF305496";
const HEADER_FONT_COLOR = "FFFFFFFF";
const LINK_COLOR = "FF0563C1";
const MIN_WIDTH = 12;
const MAX_WIDTH = 45;
const MAX_SAMPLE_LEN = 60;

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

export async function formatWorkbookFile(
  filePath: string
): Promise<{ ok: boolean; detail?: string }> {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.worksheets[0];
    if (!ws || ws.rowCount < 1 || ws.columnCount < 1) {
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
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colCount } };

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
