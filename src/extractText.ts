/**
 * Extract plain text from a resume/cover-letter file, for read_document.
 *
 * Ported from the former scripts/extract_text.py (pypdf + python-docx) to
 * pure JS (pdf-parse + mammoth) so this server has no Python dependency.
 */

import * as fs from "node:fs";
// pdf-parse@1.x is CJS (`module.exports = function pdf(buffer, opts) {...}`);
// Node's ESM/CJS interop gives us that function as the default import.
import pdfParse from "pdf-parse";
import * as mammoth from "mammoth";

export interface ExtractedText {
  text: string;
  words: number;
  chars: number;
  pages?: number;
}

/** Collapse runs of blank lines so the output stays readable — same cleanup
 *  extract_text.py did. */
function collapseBlankLines(text: string): string {
  const lines = text.split(/\r\n|\r|\n/).map((l) => l.replace(/\s+$/, ""));
  const cleaned: string[] = [];
  let blank = false;
  for (const ln of lines) {
    if (ln.trim() === "") {
      if (!blank) cleaned.push("");
      blank = true;
    } else {
      cleaned.push(ln);
      blank = false;
    }
  }
  return cleaned.join("\n").trim();
}

async function extractPdf(filePath: string): Promise<{ text: string; pages: number }> {
  const buf = fs.readFileSync(filePath);
  const data = await pdfParse(buf);
  return { text: data.text ?? "", pages: data.numpages ?? 0 };
}

async function extractDocx(filePath: string): Promise<{ text: string }> {
  const result = await mammoth.extractRawText({ path: filePath });
  return { text: result.value ?? "" };
}

/**
 * Extract plain text + metadata from a .pdf or .docx file. Throws a plain
 * Error with a clear message on an unsupported type or extraction failure;
 * the caller (read_document) wraps that into a UserFacingError.
 */
export async function extractText(filePath: string): Promise<ExtractedText> {
  const dot = filePath.lastIndexOf(".");
  const ext = dot >= 0 ? filePath.slice(dot).toLowerCase() : "";

  let raw: string;
  let pages: number | undefined;
  if (ext === ".pdf") {
    const r = await extractPdf(filePath);
    raw = r.text;
    pages = r.pages;
  } else if (ext === ".docx") {
    raw = (await extractDocx(filePath)).text;
  } else {
    throw new Error(`unsupported file type: ${ext || "(none)"}`);
  }

  const text = collapseBlankLines(raw);
  return {
    text,
    words: text ? text.split(/\s+/).filter(Boolean).length : 0,
    chars: text.length,
    ...(pages !== undefined ? { pages } : {}),
  };
}
