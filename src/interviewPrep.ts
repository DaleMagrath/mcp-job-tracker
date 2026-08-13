/**
 * Interview-prep tools (Interview_Prep_QA.md) + its markdown engine.
 */

import * as fs from "node:fs";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, wrapFsError, textResult, guard } from "./errors.js";
import { backupWithRotation } from "./workbook.js";
import { INTERVIEW_PREP_FILE, INTERVIEW_PREP_TITLE } from "./config.js";

/** Read the prep file, or null if it doesn't exist yet. */
function readInterviewPrepFile(): string | null {
  if (!fs.existsSync(INTERVIEW_PREP_FILE)) return null;
  try {
    return fs.readFileSync(INTERVIEW_PREP_FILE, "utf8");
  } catch (err) {
    throw wrapFsError(err, "read", INTERVIEW_PREP_FILE);
  }
}

function backupInterviewPrepFile(): string {
  return backupWithRotation(INTERVIEW_PREP_FILE);
}

function writeInterviewPrepFile(content: string): void {
  try {
    fs.writeFileSync(INTERVIEW_PREP_FILE, content, "utf8");
  } catch (err) {
    throw wrapFsError(err, "write", INTERVIEW_PREP_FILE);
  }
}

const H2_RE = /^##\s+/;

/**
 * Does a `## ` heading's company portion match `company` (case-insensitive)?
 * Real headings look like "## Acme Corp — Engineering Manager", so we match on
 * the part before a spaced dash as well as the whole heading text.
 */
function headingMatchesCompany(headingLine: string, company: string): boolean {
  const h = headingLine.replace(H2_RE, "").trim().toLowerCase();
  const c = company.trim().toLowerCase();
  if (!c) return false;
  if (h === c) return true;
  const firstPart = h.split(/\s+[—–-]\s+/)[0].trim();
  return firstPart === c;
}

/** All `## ` section heading texts (without the leading "## "). */
function sectionHeadings(content: string): string[] {
  return content
    .split(/\r?\n/)
    .filter((l) => H2_RE.test(l))
    .map((l) => l.replace(H2_RE, "").trim());
}

/**
 * Line indexes of every `## ` heading whose company matches. If any heading is
 * an *exact* full-title match (e.g. "Acme Corp — Booking Director"), only those are
 * returned — so a specific heading always wins over a bare prefix. Otherwise all
 * prefix matches are returned, and >1 means the caller must disambiguate.
 */
function matchingHeadingIdxs(lines: string[], company: string): number[] {
  const c = company.trim().toLowerCase();
  const idxs: number[] = [];
  for (let i = 0; i < lines.length; i++) {
    if (H2_RE.test(lines[i]) && headingMatchesCompany(lines[i], company)) {
      idxs.push(i);
    }
  }
  const exact = idxs.filter(
    (i) => lines[i].replace(H2_RE, "").trim().toLowerCase() === c
  );
  return exact.length ? exact : idxs;
}

/** Text of one section given its heading line index (through the next heading). */
function extractSectionAt(lines: string[], start: number): string {
  let end = lines.length;
  for (let i = start + 1; i < lines.length; i++) {
    if (H2_RE.test(lines[i])) {
      end = i;
      break;
    }
  }
  // Drop a trailing "---" separator / blank lines that belong between sections.
  let e = end - 1;
  while (e > start && (lines[e].trim() === "" || lines[e].trim() === "---")) e--;
  return lines.slice(start, e + 1).join("\n").trim();
}

/** The heading names matched by `company`, for building disambiguation errors. */
function ambiguousSectionNames(lines: string[], idxs: number[]): string[] {
  return idxs.map((i) => lines[i].replace(H2_RE, "").trim());
}

/**
 * Append markdown under a company's section, creating the file and/or the
 * section as needed. Returns whether a new section was created plus the backup
 * path (null when the file was created fresh). Does not itself validate inputs.
 */
function appendToInterviewPrep(
  company: string,
  body: string
): { created: boolean; backup: string | null } {
  const existing = readInterviewPrepFile();
  const fileExisted = existing !== null;
  let doc = existing ?? INTERVIEW_PREP_TITLE + "\n";

  const lines = doc.split(/\r?\n/);
  const idxs = matchingHeadingIdxs(lines, company);
  if (idxs.length > 1) {
    const names = ambiguousSectionNames(lines, idxs);
    throw new UserFacingError(
      `"${company}" matches multiple sections: ${names.join("; ")}. ` +
        `Pass the full heading (e.g. "${names[0]}") as company to disambiguate.`
    );
  }
  const headingIdx = idxs.length === 1 ? idxs[0] : -1;

  let created: boolean;
  if (headingIdx !== -1) {
    // Insert into the existing section, before its trailing "---"/blank lines.
    created = false;
    let end = lines.length;
    for (let i = headingIdx + 1; i < lines.length; i++) {
      if (H2_RE.test(lines[i])) {
        end = i;
        break;
      }
    }
    let k = end - 1;
    while (k >= headingIdx + 1 && lines[k].trim() === "") k--;
    if (k >= headingIdx + 1 && lines[k].trim() === "---") {
      k--;
      while (k >= headingIdx + 1 && lines[k].trim() === "") k--;
    }
    const insertAt = k + 1;
    lines.splice(insertAt, 0, "", ...body.split(/\r?\n/));
    doc = lines.join("\n");
  } else {
    // Create a new section; add a "---" separator if other sections exist.
    created = true;
    const hasAnySection = lines.some((l) => H2_RE.test(l));
    const base = doc.replace(/\s+$/, "");
    const sep = hasAnySection ? "\n\n---\n\n" : "\n\n";
    doc = `${base}${sep}## ${company}\n\n${body}\n`;
  }

  if (!doc.endsWith("\n")) doc += "\n";

  const backup = fileExisted ? backupInterviewPrepFile() : null;
  writeInterviewPrepFile(doc);
  return { created, backup };
}

export function register(server: McpServer): void {
  // I1. read_interview_prep -------------------------------------------
  server.registerTool(
    "read_interview_prep",
    {
      title: "Read interview prep",
      description:
        "Read the interview-prep Q&A markdown file (Interview_Prep_QA.md) and " +
        "return it as text. Pass an optional company to return just that " +
        "company's section. Returns a friendly message (not an error) if the " +
        "file doesn't exist yet or the company has no section.",
      inputSchema: {
        company: z
          .string()
          .optional()
          .describe("Return only this company's section (case-insensitive)."),
      },
    },
    async (args) =>
      guard(() => {
        const content = readInterviewPrepFile();
        if (content === null) {
          return textResult(
            "No interview prep file found yet — use append_interview_prep to create it."
          );
        }
        if (args.company && args.company.trim()) {
          const lines = content.split(/\r?\n/);
          const idxs = matchingHeadingIdxs(lines, args.company);
          if (idxs.length === 0) {
            const heads = sectionHeadings(content);
            return textResult(
              `No section found for "${args.company}". ` +
                (heads.length
                  ? `Sections in the file: ${heads.join("; ")}.`
                  : "The file has no company sections yet.")
            );
          }
          if (idxs.length > 1) {
            const names = ambiguousSectionNames(lines, idxs);
            return textResult(
              `"${args.company}" matches multiple sections: ${names.join("; ")}. ` +
                `Pass the full heading to read just one.`
            );
          }
          return textResult(extractSectionAt(lines, idxs[0]));
        }
        return textResult(content);
      })
  );

  // I2. append_interview_prep -----------------------------------------
  server.registerTool(
    "append_interview_prep",
    {
      title: "Append interview prep",
      description:
        "Append markdown to the interview-prep file under a company's section. " +
        "Creates the file (with a top-level heading) and/or the company section " +
        "if needed; otherwise adds to the existing section without duplicating " +
        "its heading. Keep the file's style: '**Q: ...**' / 'A: ...' pairs and " +
        "optional '*Notes: ...*' lines. Backs the file up before writing.",
      inputSchema: {
        company: z
          .string()
          .min(1)
          .describe("Company name; matches an existing '## Company' section or creates one."),
        content: z
          .string()
          .min(1)
          .describe(
            "Markdown to add under the company section (e.g. a Q/A pair and notes)."
          ),
      },
    },
    async (args) =>
      guard(() => {
        const company = args.company.trim();
        const body = args.content.trim();
        if (!company) throw new UserFacingError("company cannot be empty.");
        if (!body) throw new UserFacingError("content cannot be empty.");

        const { created, backup } = appendToInterviewPrep(company, body);

        return textResult({
          message: created
            ? `Created a new "${company}" section and added your content.`
            : `Appended to the existing "${company}" section.`,
          file: INTERVIEW_PREP_FILE,
          section: company,
          created,
          backup,
          added: body,
        });
      })
  );
}
