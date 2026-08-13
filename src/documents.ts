/**
 * Save & print document tools (Resumes\) + their doc/print engine.
 */

import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createRequire } from "node:module";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, wrapFsError, textResult, guard } from "./errors.js";
import {
  ensureDir,
  humanSize,
  backupWithRotation,
  openWorkbook,
  readAllRecords,
  findByCompanyPosition,
} from "./workbook.js";
import {
  RESUMES_DIR,
  RESUME_MASTER_FILE,
  SERVER_DIR,
  TRACKER,
  DISCOVERY,
  JOB_ROOT,
} from "./config.js";

const execFileAsync = promisify(execFile);
// pdf-to-printer is CJS; load its silent-print function via require interop.
const require = createRequire(import.meta.url);
const { print: sumatraPrint } = require("pdf-to-printer") as {
  print: (file: string, options?: { printer?: string }) => Promise<void>;
};

/**
 * Ceiling on the inline base64 path. Base64 is emitted token-by-token by the
 * model (~3 chars/token for binary), so a 53 KB PDF costs ~22k output tokens —
 * minutes of generation, and large files simply exceed the response budget and
 * can never finish. Files above this must come in via `source_path`, which
 * costs ~30 tokens regardless of size. ~20k chars ≈ 15 KB decoded.
 */
const MAX_INLINE_B64 = 20_000;

/** Plain-text extensions read_text_file will serve (binary stays out). */
const TEXT_EXTS = new Set([
  ".md", ".txt", ".json", ".csv", ".tsv", ".log", ".yaml", ".yml", ".xml", ".ini",
]);
/** Hard ceiling: refuse files larger than this to read at all. */
const TEXT_FILE_HARD_MAX = 5 * 1024 * 1024;
/** Soft cap: text beyond this is returned truncated, with a note. */
const TEXT_FILE_SOFT_MAX = 1_000_000;

/** Alphanumeric-only lowercasing, for loose filename↔company matching. */
function normalizeToken(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, "");
}

/**
 * Reject bytes that clearly aren't the format the extension claims. Catches a
 * truncated or mistyped base64 payload before it lands as a corrupt resume that
 * only fails much later, at print/upload time. Unknown extensions pass through.
 */
function assertMagicBytes(head: Buffer, filename: string): void {
  const ext = path.extname(filename).toLowerCase();
  const expected: Record<string, { magic: Buffer; label: string }> = {
    ".docx": { magic: Buffer.from([0x50, 0x4b, 0x03, 0x04]), label: "a .docx (ZIP)" },
    ".pdf": { magic: Buffer.from("%PDF-", "latin1"), label: "a PDF" },
  };
  const want = expected[ext];
  if (!want) return;
  if (!head.subarray(0, want.magic.length).equals(want.magic)) {
    throw new UserFacingError(
      `The content doesn't look like ${want.label} — its leading bytes are wrong. ` +
        `If it came from base64, the payload was probably truncated or altered in ` +
        `transit; prefer save_document with source_path for binary files.`
    );
  }
}

/**
 * Resume/cover-letter files in the Resumes folder whose name looks related to
 * `company` (normalized substring match). Used to link a tracker row to its
 * saved documents — e.g. get_job surfaces which resume was used for a company.
 */
export function resumeFilesFor(company: string): string[] {
  try {
    if (!fs.existsSync(RESUMES_DIR)) return [];
    const token = normalizeToken(company.trim());
    if (!token) return [];
    return fs
      .readdirSync(RESUMES_DIR)
      .filter((f) => /\.(pdf|docx)$/i.test(f))
      .filter((f) => normalizeToken(f).includes(token))
      .sort();
  } catch {
    return [];
  }
}

/** Enumerate installed printers via PowerShell (robust; avoids the buggy
 *  getPrinters in pdf-to-printer). Returns [] if enumeration fails. */
async function listPrinters(): Promise<{ name: string; isDefault: boolean }[]> {
  const psScript =
    "Get-CimInstance Win32_Printer | Select-Object Name,Default | ConvertTo-Json -Compress";
  try {
    const { stdout } = await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-Command", psScript],
      { windowsHide: true }
    );
    const text = stdout.trim();
    if (!text) return [];
    const raw = JSON.parse(text);
    const arr = Array.isArray(raw) ? raw : [raw];
    return arr.map((p: any) => ({
      name: String(p.Name),
      isDefault: p.Default === true,
    }));
  } catch {
    return [];
  }
}

/** Locate a LibreOffice soffice.exe, if installed. */
function findSoffice(): string | null {
  const candidates = [
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
    "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
  ];
  return candidates.find((p) => fs.existsSync(p)) ?? null;
}

/**
 * Convert a .docx to a .pdf next to it, so print_document always has a PDF.
 * Caches: reuses an existing PDF that is newer than the docx. Prefers
 * LibreOffice (silent headless), falls back to Microsoft Word via COM.
 */
async function convertDocxToPdf(docxPath: string): Promise<string> {
  const pdfPath = docxPath.replace(/\.docx$/i, "") + ".pdf";

  // Cache: reuse an up-to-date PDF.
  try {
    if (fs.existsSync(pdfPath)) {
      const dm = fs.statSync(docxPath).mtimeMs;
      const pm = fs.statSync(pdfPath).mtimeMs;
      if (pm >= dm) return pdfPath;
    }
  } catch {
    /* fall through and (re)convert */
  }

  // 1) LibreOffice, if present.
  const soffice = findSoffice();
  if (soffice) {
    try {
      await execFileAsync(
        soffice,
        ["--headless", "--convert-to", "pdf", "--outdir", path.dirname(pdfPath), docxPath],
        { windowsHide: true }
      );
      if (fs.existsSync(pdfPath)) return pdfPath;
    } catch {
      /* fall through to Word */
    }
  }

  // 2) Microsoft Word via COM (silent).
  const escd = docxPath.replace(/'/g, "''");
  const escp = pdfPath.replace(/'/g, "''");
  const ps = [
    "$ErrorActionPreference='Stop'",
    "$w=New-Object -ComObject Word.Application",
    "$w.Visible=$false",
    "$w.DisplayAlerts=0",
    "try{",
    `  $d=$w.Documents.Open('${escd}',$false,$true)`,
    `  $d.ExportAsFixedFormat('${escp}',17)`, // 17 = wdExportFormatPDF
    "  $d.Close(0)",
    "} finally { $w.Quit() }",
  ].join(" ");
  try {
    await execFileAsync(
      "powershell.exe",
      ["-NoProfile", "-NonInteractive", "-Command", ps],
      { windowsHide: true }
    );
  } catch (err: any) {
    throw new UserFacingError(
      "Could not convert the .docx to PDF for printing — this needs LibreOffice " +
        "or Microsoft Word installed. Alternatively, save the document as a PDF " +
        "and print that directly.\n" +
        `Details: ${err?.stderr || err?.message || String(err)}`
    );
  }
  if (!fs.existsSync(pdfPath)) {
    throw new UserFacingError(
      "docx→PDF conversion ran but produced no PDF. Try saving/printing a PDF directly."
    );
  }
  return pdfPath;
}

/** Run a bundled Python script (trying "python" then "py"), returning stdout. */
async function runPython(scriptRelPath: string, args: string[]): Promise<string> {
  const script = path.join(SERVER_DIR, "..", "scripts", scriptRelPath);
  if (!fs.existsSync(script)) {
    throw new UserFacingError(`Script not found: ${script}`);
  }
  for (const py of ["python", "py"]) {
    try {
      const { stdout } = await execFileAsync(py, [script, ...args], {
        windowsHide: true,
        maxBuffer: 32 * 1024 * 1024,
      });
      return stdout;
    } catch (err: any) {
      if (err?.code === "ENOENT") continue; // this interpreter isn't installed
      throw new UserFacingError(
        `Python script ${scriptRelPath} failed: ` +
          `${err?.stderr || err?.message || String(err)}`
      );
    }
  }
  throw new UserFacingError(
    "Python was not found. This server's document tools need Python installed " +
      "(with pypdf / python-docx for text extraction and resume generation)."
  );
}

/** Sanitize a string into a safe filename fragment. */
function safeFilePart(s: string): string {
  return s.trim().replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
}

/**
 * After generating a resume, suggest how to fold it into the tracker — returning
 * the exact follow-up tool call based on whether this company+position is already
 * tracked, a discovery lead, or brand new. This is what wires generate → apply →
 * track into one flow: the model reads `nextStep` and offers it. Best-effort;
 * falls back to an add_job suggestion if the sheets can't be read.
 */
function resumeNextStep(company: string, position: string, versionLabel: string) {
  const addJob = {
    status: "new" as const,
    suggestion:
      `"${company} — ${position}" isn't in your tracker yet. Add it as Applied ` +
      `with Resume Version "${versionLabel}"?`,
    suggestedCall: {
      tool: "add_job",
      arguments: { company, position, status: "Applied", resume_version: versionLabel },
    },
  };
  try {
    if (findByCompanyPosition(readAllRecords(openWorkbook(TRACKER)), company, position).length) {
      return {
        status: "already_tracked" as const,
        suggestion:
          `"${company} — ${position}" is already in your tracker. Set its Resume ` +
          `Version to "${versionLabel}"?`,
        suggestedCall: {
          tool: "update_job",
          arguments: { company, position, resume_version: versionLabel },
        },
      };
    }
    if (findByCompanyPosition(readAllRecords(openWorkbook(DISCOVERY)), company, position).length) {
      return {
        status: "in_discovery" as const,
        suggestion:
          `This is a discovery lead. Promote it to your tracker as Applied ` +
          `(Resume Version "${versionLabel}")?`,
        suggestedCall: {
          tool: "promote_to_tracker",
          arguments: { company, position, resume_version: versionLabel },
        },
      };
    }
    return addJob;
  } catch {
    return addJob;
  }
}

/* ---- resume_master.json field access (scoped to that one file) ---- */

/** Load and parse resume_master.json, with clear errors. */
function loadResumeMaster(): any {
  if (!fs.existsSync(RESUME_MASTER_FILE)) {
    throw new UserFacingError(
      `No resume master found at:\n  ${RESUME_MASTER_FILE}\n` +
        `Create it, or set JOB_RESUME_MASTER_FILE.`
    );
  }
  try {
    return JSON.parse(fs.readFileSync(RESUME_MASTER_FILE, "utf8"));
  } catch (err: any) {
    throw new UserFacingError(
      `resume_master.json is not valid JSON: ${err?.message || String(err)}`
    );
  }
}

/** Parse "projects[0].description" → ["projects", 0, "description"]. */
function parseJsonPath(pathStr: string): (string | number)[] {
  return pathStr
    .replace(/\[(\d+)\]/g, ".$1")
    .split(".")
    .map((p) => p.trim())
    .filter((p) => p.length > 0)
    .map((p) => (/^\d+$/.test(p) ? Number(p) : p));
}

/** Navigate to a path; reports existence plus the parent container + key. */
function resolveJsonPath(root: any, tokens: (string | number)[]) {
  let cur = root;
  for (let i = 0; i < tokens.length; i++) {
    const k = tokens[i];
    if (cur == null || typeof cur !== "object") return { found: false as const };
    if (Array.isArray(cur)) {
      if (typeof k !== "number" || k < 0 || k >= cur.length) {
        return { found: false as const };
      }
    } else if (!Object.prototype.hasOwnProperty.call(cur, k)) {
      return { found: false as const };
    }
    if (i === tokens.length - 1) {
      return { found: true as const, value: cur[k as any], parent: cur, key: k };
    }
    cur = cur[k as any];
  }
  return { found: true as const, value: root, parent: null, key: null };
}

/** Human hint about where a path went wrong (lists siblings to help). */
function pathHint(root: any, tokens: (string | number)[]): string {
  let cur = root;
  for (let i = 0; i < tokens.length; i++) {
    const k = tokens[i];
    if (cur == null || typeof cur !== "object") {
      return `"${tokens.slice(0, i).join(".")}" is not an object or array.`;
    }
    const where = tokens.slice(0, i).join(".") || "(root)";
    const container = Array.isArray(cur)
      ? `an array of length ${cur.length}`
      : `an object with keys: ${Object.keys(cur).join(", ")}`;
    const present = Array.isArray(cur)
      ? typeof k === "number" && k >= 0 && k < cur.length
      : Object.prototype.hasOwnProperty.call(cur, k);
    if (!present) return `No "${k}" at "${where}" — that level is ${container}.`;
    cur = cur[k as any];
  }
  return "";
}

/** Shallow shape of a value (keys / array length / scalar type + string length),
 *  never the full content — for list_resume_master_structure. */
function describeShape(v: any, depth = 2): any {
  if (Array.isArray(v)) {
    return {
      type: "array",
      length: v.length,
      ...(v.length && depth > 0 ? { item: describeShape(v[0], depth - 1) } : {}),
    };
  }
  if (v && typeof v === "object") {
    const keys = Object.keys(v);
    if (depth <= 0) return { type: "object", keys };
    const fields: Record<string, any> = {};
    for (const k of keys) fields[k] = describeShape(v[k], depth - 1);
    return { type: "object", fields };
  }
  if (typeof v === "string") return { type: "string", length: v.length };
  return { type: v === null ? "null" : typeof v };
}

/** Coarse category, to catch array↔object↔scalar mistakes on update. */
function valueCategory(v: any): "array" | "object" | "scalar" {
  if (Array.isArray(v)) return "array";
  if (v !== null && typeof v === "object") return "object";
  return "scalar";
}

export function register(server: McpServer): void {
  // S1. save_document -------------------------------------------------
  server.registerTool(
    "save_document",
    {
      title: "Save document",
      description:
        "Save a generated document (docx or PDF) into the Resumes folder " +
        "(created if missing).\n" +
        "PREFERRED: pass `source_path` — the full path to a file already written " +
        "to disk (e.g. one the docx/pdf skill just generated). The server copies " +
        "it; this is instant and cannot corrupt the file.\n" +
        "FALLBACK: `content_base64`, for SMALL files only (~15 KB decoded max). " +
        "Base64 must be generated token-by-token, so a typical 50 KB resume takes " +
        "minutes and may exceed the response limit outright — use `source_path` " +
        "for anything binary or non-trivial.\n" +
        "Give exactly one of the two. Refuses to overwrite an existing file " +
        "unless overwrite=true, in which case the old file is backed up first. " +
        "Returns the saved path and size.",
      inputSchema: {
        filename: z
          .string()
          .min(1)
          .describe(
            'File name only (no folders), e.g. "Dale-Magrath-Resume-Acme-Corp.docx".'
          ),
        source_path: z
          .string()
          .optional()
          .describe(
            "PREFERRED. Absolute path to an existing file to copy into the " +
              "Resumes folder. Use this for any real document."
          ),
        content_base64: z
          .string()
          .optional()
          .describe(
            "Fallback for small files only (~15 KB decoded). Base64-encoded file " +
              "bytes; a data: URL prefix is tolerated. Prefer source_path."
          ),
        overwrite: z
          .boolean()
          .optional()
          .describe("If false (default), refuse to replace an existing file."),
      },
    },
    async (args) =>
      guard(() => {
        const filename = args.filename.trim();
        if (
          !filename ||
          filename !== path.basename(filename) ||
          filename.includes("..")
        ) {
          throw new UserFacingError(
            `"${args.filename}" must be a plain file name — no folders or "..".`
          );
        }

        const sourcePath = args.source_path?.trim() || "";
        const hasB64 = !!args.content_base64?.trim();
        if (sourcePath && hasB64) {
          throw new UserFacingError(
            "Give either source_path or content_base64, not both."
          );
        }
        if (!sourcePath && !hasB64) {
          throw new UserFacingError(
            "Nothing to save — pass source_path (preferred) or content_base64."
          );
        }

        // Resolve the source: a path to copy, or decoded inline bytes.
        let buf: Buffer | null = null;
        let resolvedSource = "";
        let bytes = 0;

        if (sourcePath) {
          if (!path.isAbsolute(sourcePath)) {
            throw new UserFacingError(
              `source_path must be an absolute path (got "${sourcePath}"). The ` +
                `server's working directory is not where your files live.`
            );
          }
          resolvedSource = path.resolve(sourcePath);
          let st: fs.Stats;
          try {
            st = fs.statSync(resolvedSource);
          } catch (err) {
            throw wrapFsError(err, "read", resolvedSource, "document");
          }
          if (!st.isFile()) {
            throw new UserFacingError(`source_path is not a file:\n  ${resolvedSource}`);
          }
          if (st.size === 0) {
            throw new UserFacingError(`source_path is an empty file:\n  ${resolvedSource}`);
          }
          bytes = st.size;

          // Cheap format sanity check: read only the leading bytes.
          const head = Buffer.alloc(8);
          try {
            const fd = fs.openSync(resolvedSource, "r");
            try {
              fs.readSync(fd, head, 0, 8, 0);
            } finally {
              fs.closeSync(fd);
            }
          } catch (err) {
            throw wrapFsError(err, "read", resolvedSource, "document");
          }
          assertMagicBytes(head, filename);
        } else {
          const rawB64 = args.content_base64!;
          // Strip an optional data: URL prefix and validate the base64 payload.
          const b64 = rawB64
            .replace(/^data:[^;]*;base64,/, "")
            .replace(/\s/g, "");
          if (b64.length > MAX_INLINE_B64) {
            throw new UserFacingError(
              `content_base64 is ${b64.length} characters — too large for the ` +
                `inline path (limit ${MAX_INLINE_B64}, about 15 KB decoded). ` +
                `Write the file to disk and pass source_path instead; that is ` +
                `instant at any size.`
            );
          }
          if (!/^[A-Za-z0-9+/]*={0,2}$/.test(b64) || b64.length === 0) {
            throw new UserFacingError("content_base64 is not valid base64.");
          }
          buf = Buffer.from(b64, "base64");
          if (buf.length === 0) {
            throw new UserFacingError("Decoded content is empty — check content_base64.");
          }
          assertMagicBytes(buf.subarray(0, 8), filename);
          bytes = buf.length;
        }

        ensureDir(RESUMES_DIR);
        const target = path.join(RESUMES_DIR, filename);
        if (resolvedSource && path.resolve(target) === resolvedSource) {
          throw new UserFacingError(
            `source_path is already the destination file:\n  ${target}`
          );
        }
        const exists = fs.existsSync(target);
        if (exists && !args.overwrite) {
          throw new UserFacingError(
            `"${filename}" already exists in the Resumes folder. Pass ` +
              `overwrite: true to replace it (a .bak backup is made first).`
          );
        }

        let backup: string | null = null;
        if (exists && args.overwrite) {
          backup = backupWithRotation(target);
        }

        try {
          if (buf) fs.writeFileSync(target, buf);
          else fs.copyFileSync(resolvedSource, target);
        } catch (err) {
          throw wrapFsError(err, "write", target, "document");
        }

        return textResult({
          message: `Saved ${filename} (${humanSize(bytes)}).`,
          path: target,
          bytes,
          via: resolvedSource ? "source_path" : "content_base64",
          copiedFrom: resolvedSource || undefined,
          overwrote: exists,
          backup,
        });
      })
  );

  // S2. print_document ------------------------------------------------
  server.registerTool(
    "print_document",
    {
      title: "Print document",
      description:
        "Silently print a saved document to a printer. Give a file name in the " +
        "Resumes folder (or a full path). PDFs print directly; .docx files are " +
        "converted to PDF first (cached next to the docx). Uses the system " +
        "default printer unless printer_name is given. Confirms the job was " +
        "SENT to the printer, not that it physically finished.",
      inputSchema: {
        filename: z
          .string()
          .min(1)
          .describe("File in the Resumes folder, or a full path to a .pdf/.docx."),
        printer_name: z
          .string()
          .optional()
          .describe("Printer name to use; omit for the system default."),
      },
    },
    async (args) =>
      guard(async () => {
        const raw = args.filename.trim();
        const filePath = path.isAbsolute(raw)
          ? path.resolve(raw)
          : path.join(RESUMES_DIR, raw);
        if (!fs.existsSync(filePath)) {
          throw new UserFacingError(
            `No file found at:\n  ${filePath}\n` +
              `Save it first with save_document, or pass a valid path.`
          );
        }
        const ext = path.extname(filePath).toLowerCase();
        if (ext !== ".pdf" && ext !== ".docx") {
          throw new UserFacingError(
            `Can only print .pdf or .docx files (got "${ext || "no extension"}").`
          );
        }

        // Validate / normalise the printer choice.
        const printers = await listPrinters();
        let printerName = args.printer_name?.trim() || "";
        if (printerName) {
          const match = printers.find(
            (p) => p.name.toLowerCase() === printerName.toLowerCase()
          );
          if (!match && printers.length) {
            throw new UserFacingError(
              `No printer named "${printerName}". Available printers: ` +
                `${printers.map((p) => p.name).join(", ")}.`
            );
          }
          if (match) printerName = match.name;
        }
        const usedPrinter =
          printerName || printers.find((p) => p.isDefault)?.name || "system default";

        // Ensure we have a PDF to hand to SumatraPDF.
        let pdfPath = filePath;
        let convertedFrom: string | undefined;
        if (ext === ".docx") {
          pdfPath = await convertDocxToPdf(filePath);
          convertedFrom = filePath;
        }

        try {
          await sumatraPrint(pdfPath, printerName ? { printer: printerName } : {});
        } catch (err: any) {
          throw new UserFacingError(
            `Failed to send the print job: ${err?.message || String(err)}`
          );
        }

        return textResult({
          message: `Sent "${path.basename(filePath)}" to printer: ${usedPrinter}.`,
          note: "Confirms the job was sent to the printer, not that it physically finished printing.",
          printer: usedPrinter,
          printed: pdfPath,
          convertedFromDocx: convertedFrom,
        });
      })
  );

  // S3. list_documents ------------------------------------------------
  server.registerTool(
    "list_documents",
    {
      title: "List documents",
      description:
        "List the saved resumes / cover letters in the Resumes folder (name, " +
        "size, and modified date). Optional `filter` is a case-insensitive " +
        "substring on the file name. Newest first.",
      inputSchema: {
        filter: z
          .string()
          .optional()
          .describe("Case-insensitive substring to filter file names by."),
      },
    },
    async (args) =>
      guard(() => {
        if (!fs.existsSync(RESUMES_DIR)) {
          return textResult({ folder: RESUMES_DIR, count: 0, documents: [] });
        }
        const filter = args.filter?.trim().toLowerCase() || "";
        const docs = fs
          .readdirSync(RESUMES_DIR)
          .filter((f) => {
            const full = path.join(RESUMES_DIR, f);
            try {
              return fs.statSync(full).isFile();
            } catch {
              return false;
            }
          })
          .filter((f) => !filter || f.toLowerCase().includes(filter))
          .map((f) => {
            const st = fs.statSync(path.join(RESUMES_DIR, f));
            return {
              name: f,
              size: humanSize(st.size),
              bytes: st.size,
              modified: new Date(st.mtimeMs).toISOString().slice(0, 10),
              modifiedMs: st.mtimeMs,
            };
          })
          .sort((a, b) => b.modifiedMs - a.modifiedMs)
          .map(({ modifiedMs, ...rest }) => rest);

        return textResult({ folder: RESUMES_DIR, count: docs.length, documents: docs });
      })
  );

  // S4. delete_document -----------------------------------------------
  server.registerTool(
    "delete_document",
    {
      title: "Delete document",
      description:
        "Delete a saved document from the Resumes folder. Backs it up (to the " +
        "rotating .backups folder) first so it can be restored, then removes it. " +
        "Refuses a missing file or any path outside the Resumes folder.",
      inputSchema: {
        filename: z
          .string()
          .min(1)
          .describe('File name in the Resumes folder, e.g. "Old-Resume.pdf".'),
      },
    },
    async (args) =>
      guard(() => {
        const filename = args.filename.trim();
        if (
          !filename ||
          filename !== path.basename(filename) ||
          filename.includes("..")
        ) {
          throw new UserFacingError(
            `"${args.filename}" must be a plain file name in the Resumes folder.`
          );
        }
        const target = path.join(RESUMES_DIR, filename);
        if (!fs.existsSync(target)) {
          throw new UserFacingError(
            `No document named "${filename}" in the Resumes folder.`
          );
        }

        const backup = backupWithRotation(target);
        try {
          fs.unlinkSync(target);
        } catch (err) {
          throw wrapFsError(err, "delete", target, "document");
        }

        return textResult({
          message: `Deleted ${filename}.`,
          path: target,
          backup,
          note: "Backed up before deleting — restore from the .backups folder if needed.",
        });
      })
  );

  // S4b. read_document ------------------------------------------------
  server.registerTool(
    "read_document",
    {
      title: "Read document",
      description:
        "Extract and return the plain text of a saved resume / cover letter " +
        "(.pdf or .docx) from the Resumes folder, so it can be read and analyzed " +
        'in conversation (e.g. "read my Acme Corp resume and check it against this ' +
        'posting"). Read-only. Give a file name in the Resumes folder.',
      inputSchema: {
        filename: z
          .string()
          .min(1)
          .describe('File name in the Resumes folder, e.g. "Dale-Magrath-Resume-Acme-Corp.pdf".'),
      },
    },
    async (args) =>
      guard(async () => {
        const filename = args.filename.trim();
        if (
          !filename ||
          filename !== path.basename(filename) ||
          filename.includes("..")
        ) {
          throw new UserFacingError(
            `"${args.filename}" must be a plain file name in the Resumes folder.`
          );
        }
        const target = path.join(RESUMES_DIR, filename);
        if (!fs.existsSync(target)) {
          throw new UserFacingError(
            `No document named "${filename}" in the Resumes folder.`
          );
        }
        const ext = path.extname(target).toLowerCase();
        if (ext !== ".pdf" && ext !== ".docx") {
          throw new UserFacingError(
            `Can only read .pdf or .docx files (got "${ext || "no extension"}"). ` +
              `Supported types: .pdf, .docx.`
          );
        }

        let parsed: any;
        try {
          parsed = JSON.parse(await runPython("extract_text.py", [target]));
        } catch (err) {
          if (err instanceof UserFacingError) throw err;
          throw new UserFacingError(
            `Could not extract text from ${filename}: ${
              (err as any)?.message || String(err)
            }`
          );
        }

        // Never silently truncate — cap very long text and say so.
        const MAX = 200_000;
        let text: string = parsed.text ?? "";
        let note: string | undefined;
        if (text.length > MAX) {
          note =
            `Extracted text is ${parsed.chars} characters; returning the first ` +
            `${MAX}. Ask for a specific section if you need the rest.`;
          text = text.slice(0, MAX);
        }

        return textResult({
          filename,
          type: ext.slice(1),
          ...(parsed.pages !== undefined ? { pages: parsed.pages } : {}),
          words: parsed.words,
          chars: parsed.chars,
          ...(note ? { truncated: true, note } : {}),
          text,
        });
      })
  );

  // S5. generate_resume -----------------------------------------------
  server.registerTool(
    "generate_resume",
    {
      title: "Generate tailored resume",
      description:
        "Generate a tailored resume PDF (or docx) on the host and save it into " +
        "the Resumes folder — no base64, works from Claude Desktop.\n" +
        "The STABLE facts (employers, titles, dates, education, skills) come from " +
        "resume_master.json; YOU supply only the per-posting tailoring: a rewritten " +
        "`summary` and the `key_qualifications` bullets aligning the candidate to " +
        "this specific `company` + `position`. The server renders the master + your " +
        "tailoring into a formatted PDF (python-docx + LibreOffice) matching the " +
        "existing resumes. Do NOT invent employers, dates, or degrees — those are " +
        "fixed in the master. Returns the saved path and size, plus a `nextStep` " +
        "with a suggested follow-up call (add_job / promote_to_tracker / " +
        "update_job) — offer it to the user so tailoring, saving, and tracking " +
        "flow together.",
      inputSchema: {
        company: z.string().min(1).describe('Target company, e.g. "Acme Corp".'),
        position: z
          .string()
          .min(1)
          .describe('Target role, e.g. "Engineering Manager".'),
        summary: z
          .string()
          .min(1)
          .describe(
            "The tailored SUMMARY paragraph, rewritten to foreground the fit for " +
              "this posting. Factual — draw only on the real background."
          ),
        key_qualifications: z
          .array(z.string().min(1))
          .min(1)
          .describe(
            "Bullet points for the 'KEY QUALIFICATIONS — {company} {position} " +
              "ALIGNMENT' section, each tying real experience to this posting."
          ),
        skills: z
          .array(z.string().min(1))
          .optional()
          .describe("Optional skills list to override/reorder the master's skills."),
        include_projects: z
          .boolean()
          .optional()
          .describe("Include the Projects section (default true)."),
        include_honors: z
          .boolean()
          .optional()
          .describe(
            "Include a Honors & Awards section from the master (default false; " +
              "off keeps a full resume at 2 pages)."
          ),
        format: z
          .enum(["pdf", "docx"])
          .optional()
          .describe("Output format (default pdf)."),
        filename: z
          .string()
          .optional()
          .describe(
            "Output file name (no folders). Defaults to " +
              "Dale-Magrath-Resume-{Company}-{Position}.{ext}."
          ),
        overwrite: z
          .boolean()
          .optional()
          .describe("If false (default), refuse to replace an existing file."),
      },
    },
    async (args) =>
      guard(async () => {
        // Load the master facts.
        if (!fs.existsSync(RESUME_MASTER_FILE)) {
          throw new UserFacingError(
            `No resume master found at:\n  ${RESUME_MASTER_FILE}\n` +
              `Create it (stable facts: name, contact, experience, education, ` +
              `skills) or set JOB_RESUME_MASTER_FILE.`
          );
        }
        let master: any;
        try {
          master = JSON.parse(fs.readFileSync(RESUME_MASTER_FILE, "utf8"));
        } catch (err: any) {
          throw new UserFacingError(
            `resume_master.json is not valid JSON: ${err?.message || String(err)}`
          );
        }

        const company = args.company.trim();
        const position = args.position.trim();
        const format = args.format ?? "pdf";
        const includeProjects = args.include_projects !== false;

        // Merge master facts with the caller's per-posting tailoring.
        const spec = {
          name: master.name ?? "",
          contact: master.contact ?? "",
          summary: args.summary.trim(),
          key_qualifications: {
            target: `${company} ${position}`.trim(),
            bullets: args.key_qualifications,
          },
          experience: master.experience ?? [],
          projects: includeProjects ? master.projects ?? [] : [],
          education: master.education ?? [],
          certifications: master.certifications ?? [],
          skills: args.skills?.length ? args.skills : master.skills ?? [],
          honors: args.include_honors ? master.honors ?? [] : [],
        };

        // Resolve the output file name.
        let filename = args.filename?.trim();
        if (!filename) {
          filename = `Dale-Magrath-Resume-${safeFilePart(company)}-${safeFilePart(
            position
          )}.${format}`;
        }
        if (!filename.toLowerCase().endsWith("." + format)) {
          filename += "." + format;
        }
        if (filename !== path.basename(filename) || filename.includes("..")) {
          throw new UserFacingError(
            `"${filename}" must be a plain file name — no folders or "..".`
          );
        }

        // Render in an isolated temp dir, then copy the result into Resumes.
        const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "genresume-"));
        let produced: string;
        try {
          const specPath = path.join(tmp, "spec.json");
          const docxPath = path.join(tmp, "resume.docx");
          fs.writeFileSync(specPath, JSON.stringify(spec), "utf8");
          await runPython("gen_resume.py", [specPath, docxPath]);
          if (!fs.existsSync(docxPath)) {
            throw new UserFacingError("Generator ran but produced no .docx.");
          }
          produced = format === "pdf" ? await convertDocxToPdf(docxPath) : docxPath;

          const head = Buffer.alloc(8);
          const fd = fs.openSync(produced, "r");
          try {
            fs.readSync(fd, head, 0, 8, 0);
          } finally {
            fs.closeSync(fd);
          }
          assertMagicBytes(head, filename);

          ensureDir(RESUMES_DIR);
          const target = path.join(RESUMES_DIR, filename);
          const exists = fs.existsSync(target);
          if (exists && !args.overwrite) {
            throw new UserFacingError(
              `"${filename}" already exists in the Resumes folder. Pass ` +
                `overwrite: true to replace it (a backup is made first).`
            );
          }
          const backup = exists && args.overwrite ? backupWithRotation(target) : null;
          try {
            fs.copyFileSync(produced, target);
          } catch (err) {
            throw wrapFsError(err, "write", target, "document");
          }
          const bytes = fs.statSync(target).size;
          const versionLabel = filename
            .replace(/^Dale-Magrath-Resume-/i, "")
            .replace(/\.(pdf|docx)$/i, "");

          return textResult({
            message: `Generated ${filename} (${humanSize(bytes)}) for ${company} — ${position}.`,
            path: target,
            bytes,
            format,
            company,
            position,
            overwrote: exists,
            backup,
            nextStep: resumeNextStep(company, position, versionLabel),
          });
        } finally {
          try {
            fs.rmSync(tmp, { recursive: true, force: true });
          } catch {
            /* best-effort temp cleanup */
          }
        }
      })
  );

  // S6. get_resume_master_field ---------------------------------------
  server.registerTool(
    "get_resume_master_field",
    {
      title: "Get resume master field",
      description:
        "Read one field from resume_master.json (the stable resume facts) by " +
        'dot/bracket path, e.g. "skills", "projects[0].description", ' +
        '"experience[2].title". Scoped to that one file. Use ' +
        "list_resume_master_structure first if you're unsure of the path.",
      inputSchema: {
        path: z
          .string()
          .min(1)
          .describe('Dot/bracket path, e.g. "experience[0].dates".'),
      },
    },
    async (args) =>
      guard(() => {
        const root = loadResumeMaster();
        const tokens = parseJsonPath(args.path);
        if (!tokens.length) {
          throw new UserFacingError('Provide a path, e.g. "skills".');
        }
        const r = resolveJsonPath(root, tokens);
        if (!r.found) {
          throw new UserFacingError(
            `Path "${args.path}" not found. ${pathHint(root, tokens)} ` +
              `Use list_resume_master_structure to explore.`
          );
        }
        return textResult({ path: args.path, value: r.value });
      })
  );

  // S7. update_resume_master_field ------------------------------------
  server.registerTool(
    "update_resume_master_field",
    {
      title: "Update resume master field",
      description:
        "Update one existing field in resume_master.json by dot/bracket path " +
        "(e.g. a corrected date, an added skill, a new project). Only edits " +
        "fields that ALREADY exist — it will not create new keys (that's a " +
        "deliberate schema change to make by hand). Backs the file up first and " +
        "reports old → new. Scoped to resume_master.json only.",
      inputSchema: {
        path: z
          .string()
          .min(1)
          .describe('Existing path to update, e.g. "skills" or "projects[0].description".'),
        value: z
          .any()
          .describe("New value — string, number, boolean, array, or object."),
      },
    },
    async (args) =>
      guard(() => {
        const root = loadResumeMaster();
        const tokens = parseJsonPath(args.path);
        if (!tokens.length) {
          throw new UserFacingError("Provide a path to an existing field.");
        }
        const r = resolveJsonPath(root, tokens);
        if (!r.found || r.parent == null || r.key == null) {
          throw new UserFacingError(
            `Path "${args.path}" doesn't exist, so there's nothing to update. ` +
              `${pathHint(root, tokens)} This tool only edits existing fields — it ` +
              `won't create new keys. Use list_resume_master_structure to find the ` +
              `right path.`
          );
        }

        const oldValue = r.value;
        const oc = valueCategory(oldValue);
        const nc = valueCategory(args.value);
        if (oc !== nc) {
          throw new UserFacingError(
            `Type mismatch at "${args.path}": it currently holds ${
              oc === "scalar" ? `a ${typeof oldValue}` : `an ${oc}`
            }, but the new value is ${
              nc === "scalar" ? `a ${typeof args.value}` : `an ${nc}`
            }. That's usually a mistake — this tool won't change a field's kind ` +
              `(array/object/scalar).`
          );
        }

        (r.parent as any)[r.key as any] = args.value;
        const backup = backupWithRotation(RESUME_MASTER_FILE);
        try {
          fs.writeFileSync(
            RESUME_MASTER_FILE,
            JSON.stringify(root, null, 2) + "\n",
            "utf8"
          );
        } catch (err) {
          throw wrapFsError(err, "write", RESUME_MASTER_FILE, "document");
        }

        return textResult({
          message: `Updated "${args.path}".`,
          path: args.path,
          from: oldValue,
          to: args.value,
          backup,
        });
      })
  );

  // S8. list_resume_master_structure ----------------------------------
  server.registerTool(
    "list_resume_master_structure",
    {
      title: "List resume master structure",
      description:
        "Show the SHAPE of resume_master.json (keys, array lengths, scalar types " +
        "— not the full field values) to help pick a path for " +
        "get_resume_master_field / update_resume_master_field. Optional `path` " +
        "scopes the listing to a sub-section.",
      inputSchema: {
        path: z
          .string()
          .optional()
          .describe('Optional sub-path to scope the listing, e.g. "experience".'),
      },
    },
    async (args) =>
      guard(() => {
        const root = loadResumeMaster();
        let target = root;
        let scope = "(root)";
        if (args.path && args.path.trim()) {
          const tokens = parseJsonPath(args.path);
          const r = resolveJsonPath(root, tokens);
          if (!r.found) {
            throw new UserFacingError(
              `Path "${args.path}" not found. ${pathHint(root, tokens)}`
            );
          }
          target = r.value;
          scope = args.path;
        }
        return textResult({ path: scope, structure: describeShape(target, 2) });
      })
  );

  // S9. read_text_file ------------------------------------------------
  server.registerTool(
    "read_text_file",
    {
      title: "Read text file",
      description:
        "Read any plain-text file (.md, .txt, .json, .csv, …) inside the Job " +
        "Tracking folder tree — handoff notes, scratch files, etc. Give a " +
        "filename or relative path (e.g. \"HANDOFF_2026-08-02.md\" or " +
        '"notes/plan.txt"), resolved under the Job Tracking root. Read-only, and ' +
        "strictly scoped to that tree — paths that escape it are refused. For " +
        "PDFs/docx in the Resumes folder use read_document instead.",
      inputSchema: {
        filename: z
          .string()
          .min(1)
          .describe(
            'Filename or relative path under the Job Tracking root, e.g. "notes/plan.txt".'
          ),
      },
    },
    async (args) =>
      guard(() => {
        const raw = args.filename.trim();
        if (!raw) throw new UserFacingError("Provide a filename.");

        // Resolve under the root and enforce containment (blocks .. and abs paths).
        const resolved = path.resolve(JOB_ROOT, raw);
        const rel = path.relative(JOB_ROOT, resolved);
        if (rel === "" || rel.startsWith("..") || path.isAbsolute(rel)) {
          throw new UserFacingError(
            `"${raw}" is outside the Job Tracking folder. read_text_file can only ` +
              `read files inside:\n  ${JOB_ROOT}`
          );
        }

        const ext = path.extname(resolved).toLowerCase();
        if (!TEXT_EXTS.has(ext)) {
          if (ext === ".pdf" || ext === ".docx") {
            throw new UserFacingError(
              `"${raw}" is a ${ext} — use read_document (it extracts PDF/docx text ` +
                `from the Resumes folder).`
            );
          }
          throw new UserFacingError(
            `Can't read "${ext || "no extension"}" — read_text_file only serves ` +
              `plain text: ${[...TEXT_EXTS].join(", ")}.`
          );
        }

        if (!fs.existsSync(resolved)) {
          // List what's actually in the target directory to catch typos.
          const dir = path.dirname(resolved);
          const relDir = path.relative(JOB_ROOT, dir) || ".";
          let listing: string[] = [];
          try {
            if (fs.existsSync(dir)) {
              listing = fs.readdirSync(dir, { withFileTypes: true }).map((e) =>
                e.isDirectory() ? e.name + "/" : e.name
              );
            }
          } catch {
            /* ignore listing failure */
          }
          throw new UserFacingError(
            `No file at "${raw}". ` +
              (listing.length
                ? `Files in "${relDir}": ${listing.join(", ")}.`
                : `The directory "${relDir}" doesn't exist under the Job Tracking root.`)
          );
        }

        let st: fs.Stats;
        try {
          st = fs.statSync(resolved);
        } catch (err) {
          throw wrapFsError(err, "read", resolved, "document");
        }
        if (st.isDirectory()) {
          throw new UserFacingError(`"${raw}" is a directory, not a file.`);
        }
        if (st.size > TEXT_FILE_HARD_MAX) {
          throw new UserFacingError(
            `"${raw}" is ${humanSize(st.size)} — too large to read (limit ` +
              `${humanSize(TEXT_FILE_HARD_MAX)}).`
          );
        }

        let content: string;
        try {
          content = fs.readFileSync(resolved, "utf8");
        } catch (err) {
          throw wrapFsError(err, "read", resolved, "document");
        }

        let note: string | undefined;
        if (content.length > TEXT_FILE_SOFT_MAX) {
          note =
            `File is ${content.length} characters; returning the first ` +
            `${TEXT_FILE_SOFT_MAX}. Ask for a specific part if you need more.`;
          content = content.slice(0, TEXT_FILE_SOFT_MAX);
        }

        return textResult({
          path: rel.split(path.sep).join("/"),
          bytes: st.size,
          chars: content.length,
          ...(note ? { truncated: true, note } : {}),
          content,
        });
      })
  );
}
