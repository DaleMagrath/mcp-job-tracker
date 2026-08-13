/**
 * Tracker tools (Job_Tracking.xlsx) + column-structure tools.
 */

import * as XLSX from "xlsx";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import {
  serialToISO,
  isoToSerial,
  todaySerial,
  todayISO,
  daysSince,
} from "./dates.js";
import {
  COLUMNS,
  Column,
  DEFAULT_STATUS,
  KNOWN_STATUSES,
  isDateColumn,
} from "./config.js";
import {
  openWorkbook,
  readAllRecords,
  findByCompanyPosition,
  present,
  eq,
  includesCI,
  writeCell,
  backupFile,
  saveWorkbook,
  growRange,
  deleteRow,
  assertNotCanonical,
  deleteColumn,
  findHeader,
  readHeaderRow,
  resolveExtraFields,
} from "./workbook.js";
import { resumeFilesFor } from "./documents.js";

export function register(server: McpServer): void {
// 1. list_jobs -------------------------------------------------------
server.registerTool(
  "list_jobs",
  {
    title: "List jobs",
    description:
      "List all tracked job applications, with optional filters. Filter by " +
      "status (case-insensitive exact match), company (case-insensitive " +
      "substring), and/or a Date Applied range (inclusive, YYYY-MM-DD).",
    inputSchema: {
      status: z
        .string()
        .optional()
        .describe(`Exact status to filter by, e.g. one of: ${KNOWN_STATUSES.join(", ")}`),
      company: z
        .string()
        .optional()
        .describe("Company name substring to filter by."),
      applied_from: z
        .string()
        .optional()
        .describe("Only include jobs applied on/after this date (YYYY-MM-DD)."),
      applied_to: z
        .string()
        .optional()
        .describe("Only include jobs applied on/before this date (YYYY-MM-DD)."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      let records = readAllRecords(h);

      if (args.status) {
        records = records.filter((r) => eq(r.values["Status"], args.status!));
      }
      if (args.company) {
        records = records.filter((r) =>
          includesCI(r.values["Company"], args.company!)
        );
      }
      if (args.applied_from) {
        const from = isoToSerial(args.applied_from);
        records = records.filter((r) => {
          const s = r.serials["Date Applied"];
          return s !== undefined && s >= from;
        });
      }
      if (args.applied_to) {
        const to = isoToSerial(args.applied_to);
        records = records.filter((r) => {
          const s = r.serials["Date Applied"];
          return s !== undefined && s <= to;
        });
      }

      return textResult({
        file: h.spec.filePath,
        count: records.length,
        jobs: records.map((r) => present(h, r)),
      });
    })
);

// 2. get_job ---------------------------------------------------------
server.registerTool(
  "get_job",
  {
    title: "Get job",
    description:
      "Return full details for a single application, matched by company " +
      "(and position, if the company has more than one row).",
    inputSchema: {
      company: z.string().describe("Company name (case-insensitive)."),
      position: z
        .string()
        .optional()
        .describe("Position/title, needed only if the company has multiple rows."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const matches = findByCompanyPosition(records, args.company, args.position);

      if (matches.length === 0) {
        throw new UserFacingError(
          `No job found for company "${args.company}"` +
            (args.position ? ` and position "${args.position}"` : "") +
            "."
        );
      }
      if (matches.length > 1) {
        return textResult({
          note:
            `Multiple rows match "${args.company}". Specify a position to ` +
            `narrow it down.`,
          matches: matches.map((r) =>
            present(h, r, {
              daysSinceApplied:
                r.serials["Date Applied"] !== undefined
                  ? daysSince(r.serials["Date Applied"]!)
                  : null,
            })
          ),
        });
      }

      const r = matches[0];
      const resumeFiles = resumeFilesFor(r.values["Company"]);
      return textResult(
        present(h, r, {
          daysSinceApplied:
            r.serials["Date Applied"] !== undefined
              ? daysSince(r.serials["Date Applied"]!)
              : null,
          ...(resumeFiles.length ? { resumeFiles } : {}),
        })
      );
    })
);

// 3. add_job ---------------------------------------------------------
server.registerTool(
  "add_job",
  {
    title: "Add job",
    description:
      "Append a new job application row. Company and Position are required; " +
      "all other fields are optional. Status defaults to " +
      `"${DEFAULT_STATUS}" and Date Applied defaults to today if omitted. ` +
      "Use extra_fields to set any custom columns you've added. " +
      "Backs the file up to <file>.bak before writing.",
    inputSchema: {
      company: z.string().min(1).describe("Company name (required)."),
      position: z.string().min(1).describe("Position / title (required)."),
      job_link: z.string().optional().describe("URL of the job posting."),
      location: z.string().optional(),
      resume_version: z.string().optional().describe("Which resume version was used."),
      contact_referral: z.string().optional().describe("Contact or referral name."),
      date_applied: z
        .string()
        .optional()
        .describe("Date applied (YYYY-MM-DD). Defaults to today."),
      status: z
        .string()
        .optional()
        .describe(`Status. Defaults to "${DEFAULT_STATUS}".`),
      next_follow_up: z
        .string()
        .optional()
        .describe("Next follow-up date (YYYY-MM-DD)."),
      notes: z.string().optional(),
      extra_fields: z
        .record(z.string(), z.string())
        .optional()
        .describe(
          'Values for custom columns you\'ve added, as {"Column Name": "value"}. ' +
            "Each name must be an existing column (create it first with add_column)."
        ),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);

      const values: Record<Column, string> = {
        Company: args.company,
        Position: args.position,
        "Job Link": args.job_link ?? "",
        Location: args.location ?? "",
        "Resume Version": args.resume_version ?? "",
        "Contact/Referral": args.contact_referral ?? "",
        "Date Applied": args.date_applied ?? todayISO(),
        Status: args.status ?? DEFAULT_STATUS,
        "Next Follow-Up": args.next_follow_up ?? "",
        Notes: args.notes ?? "",
      };

      // Resolve any custom-column values, rejecting unknown/base-column names.
      const extras = resolveExtraFields(h, args.extra_fields);

      // Validate any provided dates up front (throws UserFacingError on bad input).
      if (values["Date Applied"]) isoToSerial(values["Date Applied"]);
      if (values["Next Follow-Up"]) isoToSerial(values["Next Follow-Up"]);

      // The new row goes right after the last non-blank record (or the header).
      const templateRow =
        records.length > 0 ? records[records.length - 1].row : h.firstDataRow - 1;
      const newRow =
        records.length > 0 ? records[records.length - 1].row + 1 : h.firstDataRow;

      const backup = backupFile(h);
      for (const col of COLUMNS) {
        writeCell(h, newRow, col, values[col], templateRow);
      }
      for (const e of extras) {
        writeCell(h, newRow, e.col, e.value, templateRow);
      }
      growRange(h, newRow);
      saveWorkbook(h);

      const added: Record<string, string> = Object.fromEntries(
        COLUMNS.filter((c) => values[c] !== "").map((c) => [c, values[c]])
      );
      for (const e of extras) if (e.value !== "") added[e.col] = e.value;

      return textResult({
        message: `Added ${values.Company} — ${values.Position}.`,
        backup,
        added,
      });
    })
);

// 4. update_job_status ----------------------------------------------
server.registerTool(
  "update_job_status",
  {
    title: "Update job status",
    description:
      "Update the Status (and optionally the Notes) of an existing application, " +
      "matched by company + position. Validates the row exists and backs the " +
      "file up before writing.",
    inputSchema: {
      company: z.string().min(1).describe("Company name (case-insensitive)."),
      position: z.string().min(1).describe("Position / title (case-insensitive)."),
      status: z
        .string()
        .min(1)
        .describe(`New status, e.g. one of: ${KNOWN_STATUSES.join(", ")}`),
      notes: z
        .string()
        .optional()
        .describe("Optional new Notes value (replaces existing Notes)."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const matches = findByCompanyPosition(records, args.company, args.position);

      if (matches.length === 0) {
        throw new UserFacingError(
          `No row found for "${args.company}" — "${args.position}". ` +
            `Nothing was changed. Use list_jobs to see exact company/position values.`
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `${matches.length} rows match "${args.company}" — "${args.position}". ` +
            `Refusing to update an ambiguous match; please disambiguate.`
        );
      }

      const rec = matches[0];
      const prevStatus = rec.values["Status"];

      const backup = backupFile(h);
      writeCell(h, rec.row, "Status", args.status, rec.row);
      if (args.notes !== undefined) {
        writeCell(h, rec.row, "Notes", args.notes, rec.row);
      }
      saveWorkbook(h);

      return textResult({
        message: `Updated ${rec.values.Company} — ${rec.values.Position}: status "${prevStatus}" → "${args.status}".`,
        notesUpdated: args.notes !== undefined,
        backup,
      });
    })
);

// 4b. delete_job -----------------------------------------------------
server.registerTool(
  "delete_job",
  {
    title: "Delete job",
    description:
      "Permanently remove a job application row, matched by company + position, " +
      "shifting the rows below it up (like Excel's Delete Row). Refuses to act " +
      "on a missing or ambiguous match, backs the file up to <file>.bak first, " +
      "and returns the full deleted row so it can be re-added if needed.",
    inputSchema: {
      company: z.string().min(1).describe("Company name (case-insensitive)."),
      position: z
        .string()
        .min(1)
        .describe(
          "Position / title (case-insensitive). Required so the correct row is " +
            "removed even when a company has several entries."
        ),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const matches = findByCompanyPosition(records, args.company, args.position);

      if (matches.length === 0) {
        throw new UserFacingError(
          `No row found for "${args.company}" — "${args.position}". ` +
            `Nothing was deleted. Use list_jobs to see exact company/position values.`
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `${matches.length} rows match "${args.company}" — "${args.position}". ` +
            `Refusing to delete an ambiguous match; please disambiguate.`
        );
      }

      const rec = matches[0];
      const deleted = present(h, rec); // capture before removing, so it's recoverable

      const backup = backupFile(h);
      deleteRow(h, rec.row);
      saveWorkbook(h);

      return textResult({
        message: `Deleted ${rec.values.Company} — ${rec.values.Position}.`,
        backup,
        deletedRow: deleted,
        note: "Backed up before deleting; use add_job with the fields above to restore.",
      });
    })
);

// 4c. update_job (edit any field) -----------------------------------
/** Maps update_job's optional arguments to their spreadsheet columns. */
const UPDATE_FIELD_MAP: Record<string, Column> = {
  new_company: "Company",
  new_position: "Position",
  job_link: "Job Link",
  location: "Location",
  resume_version: "Resume Version",
  contact_referral: "Contact/Referral",
  date_applied: "Date Applied",
  status: "Status",
  next_follow_up: "Next Follow-Up",
  notes: "Notes",
};

server.registerTool(
  "update_job",
  {
    title: "Update job (any field)",
    description:
      "Edit any field(s) of an existing application, matched by company + " +
      "position. Only the fields you supply are changed; pass an empty string " +
      "to clear a field. Rename with new_company / new_position. Set custom " +
      "columns via extra_fields. Refuses missing or ambiguous matches, validates " +
      "dates, and backs the file up before writing.",
    inputSchema: {
      company: z.string().min(1).describe("Current company name (case-insensitive)."),
      position: z
        .string()
        .min(1)
        .describe("Current position / title (case-insensitive)."),
      new_company: z.string().optional().describe("Rename the company."),
      new_position: z.string().optional().describe("Rename the position / title."),
      job_link: z.string().optional().describe("Job posting URL."),
      location: z.string().optional(),
      resume_version: z.string().optional(),
      contact_referral: z.string().optional(),
      date_applied: z
        .string()
        .optional()
        .describe("Date applied (YYYY-MM-DD, or empty string to clear)."),
      status: z
        .string()
        .optional()
        .describe(`Status, e.g. one of: ${KNOWN_STATUSES.join(", ")}`),
      next_follow_up: z
        .string()
        .optional()
        .describe("Next follow-up date (YYYY-MM-DD, or empty string to clear)."),
      notes: z.string().optional(),
      extra_fields: z
        .record(z.string(), z.string())
        .optional()
        .describe(
          'Custom columns to set, as {"Column Name": "value"}. Empty string ' +
            "clears a column. Each name must be an existing custom column."
        ),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const matches = findByCompanyPosition(records, args.company, args.position);

      if (matches.length === 0) {
        throw new UserFacingError(
          `No row found for "${args.company}" — "${args.position}". ` +
            `Nothing was changed. Use list_jobs to see exact company/position values.`
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `${matches.length} rows match "${args.company}" — "${args.position}". ` +
            `Refusing to update an ambiguous match; please disambiguate.`
        );
      }
      const rec = matches[0];

      // Collect only the fields the caller actually supplied.
      const argRecord = args as Record<string, unknown>;
      const updates: { col: string; value: string }[] = [];
      for (const [arg, col] of Object.entries(UPDATE_FIELD_MAP)) {
        if (argRecord[arg] !== undefined) {
          updates.push({ col, value: argRecord[arg] as string });
        }
      }
      // Custom columns, validated against the sheet's actual columns.
      for (const e of resolveExtraFields(h, args.extra_fields)) {
        updates.push(e);
      }
      if (updates.length === 0) {
        throw new UserFacingError(
          "No fields to update were provided. Supply at least one field " +
            "(e.g. status, notes, location, new_company, extra_fields, …)."
        );
      }

      // Validate any supplied dates up front, before touching the file.
      for (const u of updates) {
        if (isDateColumn(h.spec, u.col) && u.value !== "") isoToSerial(u.value);
      }

      const changed = updates.map((u) => ({
        field: u.col,
        from: rec.values[u.col],
        to: u.value,
      }));

      const backup = backupFile(h);
      for (const u of updates) {
        writeCell(h, rec.row, u.col, u.value, rec.row);
      }
      saveWorkbook(h);

      // Re-read so the returned record reflects formatted dates etc.
      const h2 = openWorkbook();
      const after = readAllRecords(h2).find((r) => r.row === rec.row);

      return textResult({
        message: `Updated ${rec.values.Company} — ${rec.values.Position} (${changed.length} field(s)).`,
        changed,
        job: after ? present(h2, after) : undefined,
        backup,
      });
    })
);

// 5. get_stale_jobs --------------------------------------------------
server.registerTool(
  "get_stale_jobs",
  {
    title: "Get stale jobs",
    description:
      'List applications still "Awaiting Response" or "Applied" whose Date ' +
      "Applied is more than N days ago (default 14), computed against today.",
    inputSchema: {
      days: z
        .number()
        .int()
        .positive()
        .optional()
        .describe("Threshold in days (default 14). Returns jobs older than this."),
    },
  },
  async (args) =>
    guard(() => {
      const threshold = args.days ?? 14;
      const h = openWorkbook();
      const records = readAllRecords(h);
      const stale = records
        .filter((r) => {
          const st = r.values["Status"];
          const isOpen = eq(st, "Awaiting Response") || eq(st, "Applied");
          if (!isOpen) return false;
          const s = r.serials["Date Applied"];
          return s !== undefined && daysSince(s) > threshold;
        })
        .map((r) => ({
          ...present(h, r),
          daysSinceApplied: daysSince(r.serials["Date Applied"]!),
        }))
        .sort((a, b) => (b.daysSinceApplied as number) - (a.daysSinceApplied as number));

      return textResult({
        today: todayISO(),
        thresholdDays: threshold,
        count: stale.length,
        staleJobs: stale,
      });
    })
);

// 5b. get_due_followups ---------------------------------------------
server.registerTool(
  "get_due_followups",
  {
    title: "Get due follow-ups",
    description:
      "List applications whose Next Follow-Up date is due — on or before today " +
      "(or within the next `days` days). Skips clearly-closed rows (Declined / " +
      "Withdrawn / Closed). Sorted most-overdue first; daysUntil is negative " +
      "when overdue.",
    inputSchema: {
      days: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          "Look-ahead window in days (default 0 = only due today or overdue)."
        ),
    },
  },
  async (args) =>
    guard(() => {
      const ahead = args.days ?? 0;
      const cutoff = todaySerial() + ahead;
      const closed = new Set([
        "declined",
        "withdrawn",
        "closed - no longer available",
      ]);
      const h = openWorkbook();
      const records = readAllRecords(h);
      const due = records
        .filter((r) => {
          if (closed.has((r.values["Status"] ?? "").trim().toLowerCase())) {
            return false;
          }
          const s = r.serials["Next Follow-Up"];
          return s !== undefined && s <= cutoff;
        })
        .map((r) => ({
          ...present(h, r),
          daysUntil: r.serials["Next Follow-Up"]! - todaySerial(),
        }))
        .sort((a, b) => (a.daysUntil as number) - (b.daysUntil as number));

      return textResult({
        today: todayISO(),
        lookAheadDays: ahead,
        count: due.length,
        dueFollowUps: due,
      });
    })
);

// 6. search_jobs -----------------------------------------------------
server.registerTool(
  "search_jobs",
  {
    title: "Search jobs",
    description:
      "Case-insensitive substring search across the Company, Position, and " +
      "Notes columns.",
    inputSchema: {
      query: z.string().min(1).describe("Keyword or phrase to search for."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const q = args.query;
      const hits = records.filter(
        (r) =>
          includesCI(r.values["Company"], q) ||
          includesCI(r.values["Position"], q) ||
          includesCI(r.values["Notes"], q)
      );
      return textResult({
        query: q,
        count: hits.length,
        jobs: hits.map((r) => present(h, r)),
      });
    })
);

// Stretch: draft_followup_message -----------------------------------
server.registerTool(
  "draft_followup_message",
  {
    title: "Draft follow-up message",
    description:
      "Draft a short, professional follow-up email for a given application, " +
      "referencing how long it has been since applying, the role, and the " +
      "resume version used. Returns draft text only — it does not send anything.",
    inputSchema: {
      company: z.string().min(1).describe("Company name (case-insensitive)."),
      position: z
        .string()
        .optional()
        .describe("Position, if the company has more than one row."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const records = readAllRecords(h);
      const matches = findByCompanyPosition(records, args.company, args.position);

      if (matches.length === 0) {
        throw new UserFacingError(
          `No job found for "${args.company}"` +
            (args.position ? ` — "${args.position}"` : "") +
            "."
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `Multiple rows match "${args.company}". Specify a position.`
        );
      }

      const r = matches[0];
      const company = r.values["Company"];
      const position = r.values["Position"];
      const contact = r.values["Contact/Referral"];
      const resume = r.values["Resume Version"];
      const appliedSerial = r.serials["Date Applied"];
      const days =
        appliedSerial !== undefined ? daysSince(appliedSerial) : undefined;
      const appliedISO =
        appliedSerial !== undefined ? serialToISO(appliedSerial) : undefined;

      const greeting = contact ? `Hi ${contact.split(/[ ,]/)[0]},` : "Hello,";
      const timePhrase =
        days !== undefined
          ? days >= 14
            ? `It has been about ${Math.round(days / 7)} weeks (${days} days) since I applied`
            : `It has been ${days} days since I applied`
          : "I recently applied";
      const resumeLine = resume
        ? ` I applied using my ${resume} resume.`
        : "";

      const subject = `Following up — ${position} application`;
      const body =
        `${greeting}\n\n` +
        `I wanted to follow up on my application for the ${position} role at ${company}. ` +
        `${timePhrase}${appliedISO ? ` (on ${appliedISO})` : ""}, and I remain very ` +
        `interested in the opportunity.${resumeLine}\n\n` +
        `Is there any update on the status of my application, or anything further ` +
        `you need from me? I'd welcome the chance to discuss how I can contribute.\n\n` +
        `Thank you for your time and consideration.\n\n` +
        `Best regards`;

      return textResult({
        note: "Draft only — review and send yourself.",
        to: contact || null,
        daysSinceApplied: days ?? null,
        subject,
        body,
      });
    })
);

// 7. add_column -----------------------------------------------------
server.registerTool(
  "add_column",
  {
    title: "Add column",
    description:
      "Add a new custom column to the spreadsheet, appended after the last " +
      "column. Optionally fill every existing job row with a default value. " +
      "Refuses if a column with that name already exists. Backs the file up " +
      "before writing. (Note: the built-in job tools only read/write the " +
      "standard columns; custom columns are managed with these column tools.)",
    inputSchema: {
      name: z.string().min(1).describe("Header text for the new column."),
      default_value: z
        .string()
        .optional()
        .describe("If given, written into this column for every existing job row."),
    },
  },
  async (args) =>
    guard(() => {
      const name = args.name.trim();
      if (!name) throw new UserFacingError("Column name cannot be empty.");

      const h = openWorkbook();
      if (findHeader(h, name)) {
        throw new UserFacingError(
          `A column named "${name}" already exists. Column names must be unique.`
        );
      }
      const records = readAllRecords(h);

      const backup = backupFile(h);
      const rng = XLSX.utils.decode_range(h.ws["!ref"]!);
      const headerRow = rng.s.r;
      const newCol = rng.e.c + 1;

      // Write the header, carrying over the neighbouring header's style so it
      // matches the existing header formatting (bold, fill, etc. where kept).
      const tmpl = h.ws[
        XLSX.utils.encode_cell({ r: headerRow, c: newCol - 1 })
      ] as XLSX.CellObject | undefined;
      const headerCell: XLSX.CellObject = { t: "s", v: name };
      if (tmpl && (tmpl as any).s) (headerCell as any).s = (tmpl as any).s;
      h.ws[XLSX.utils.encode_cell({ r: headerRow, c: newCol })] = headerCell;

      // Optionally seed every existing job row with a default value.
      const fill = args.default_value;
      if (fill !== undefined && fill !== "") {
        for (const rec of records) {
          h.ws[XLSX.utils.encode_cell({ r: rec.row, c: newCol })] = {
            t: "s",
            v: String(fill),
          };
        }
      }

      // Grow the range to include the new column.
      rng.e.c = newCol;
      h.ws["!ref"] = XLSX.utils.encode_range(rng);

      saveWorkbook(h);

      return textResult({
        message: `Added column "${name}"${
          fill !== undefined && fill !== ""
            ? ` and set it to "${fill}" on ${records.length} row(s)`
            : ""
        }.`,
        backup,
      });
    })
);

// 8. rename_column / modify_column ----------------------------------
server.registerTool(
  "modify_column",
  {
    title: "Modify column",
    description:
      "Modify a custom column: rename it (new_name) and/or set the same value " +
      "into every existing job row (fill_value; pass an empty string to clear " +
      "the column). At least one of new_name / fill_value is required. Refuses " +
      "on the built-in columns and on a missing column. Backs the file up first.",
    inputSchema: {
      name: z.string().min(1).describe("Current header text of the column to modify."),
      new_name: z.string().optional().describe("New header text (renames the column)."),
      fill_value: z
        .string()
        .optional()
        .describe(
          "If given, written into this column for every existing job row " +
            "(empty string clears the column)."
        ),
    },
  },
  async (args) =>
    guard(() => {
      if (args.new_name === undefined && args.fill_value === undefined) {
        throw new UserFacingError(
          "Nothing to do: supply new_name (to rename) and/or fill_value (to set values)."
        );
      }

      const h = openWorkbook();
      const target = findHeader(h, args.name);
      if (!target) {
        const names = readHeaderRow(h)
          .filter((x) => x.name)
          .map((x) => x.name);
        throw new UserFacingError(
          `No column named "${args.name}". Existing columns: ${names.join(", ")}.`
        );
      }
      assertNotCanonical(target.name, "modified");

      const rng = XLSX.utils.decode_range(h.ws["!ref"]!);
      const headerRow = rng.s.r;
      const changes: string[] = [];

      // Rename: validate the new name is non-empty and not already taken.
      if (args.new_name !== undefined) {
        const nn = args.new_name.trim();
        if (!nn) throw new UserFacingError("new_name cannot be empty.");
        const clash = findHeader(h, nn);
        if (clash && clash.col !== target.col) {
          throw new UserFacingError(
            `A column named "${nn}" already exists. Column names must be unique.`
          );
        }
        assertNotCanonical(nn, "used as a rename target");
      }

      const records = readAllRecords(h);
      const backup = backupFile(h);

      if (args.new_name !== undefined) {
        const nn = args.new_name.trim();
        const addr = XLSX.utils.encode_cell({ r: headerRow, c: target.col });
        const cell = (h.ws[addr] as XLSX.CellObject) ?? { t: "s", v: nn };
        cell.t = "s";
        cell.v = nn;
        delete (cell as any).w; // drop any cached formatted text
        h.ws[addr] = cell;
        changes.push(`renamed "${target.name}" → "${nn}"`);
      }

      if (args.fill_value !== undefined) {
        for (const rec of records) {
          const addr = XLSX.utils.encode_cell({ r: rec.row, c: target.col });
          if (args.fill_value === "") delete h.ws[addr];
          else h.ws[addr] = { t: "s", v: String(args.fill_value) };
        }
        changes.push(
          args.fill_value === ""
            ? `cleared the column on ${records.length} row(s)`
            : `set the column to "${args.fill_value}" on ${records.length} row(s)`
        );
      }

      saveWorkbook(h);

      return textResult({
        message: `Modified column: ${changes.join("; ")}.`,
        backup,
      });
    })
);

// 9. delete_column --------------------------------------------------
server.registerTool(
  "delete_column",
  {
    title: "Delete column",
    description:
      "Permanently remove a custom column and shift the columns to its right " +
      "left by one (like Excel's Delete Column). Refuses on the built-in " +
      "columns and on a missing column. Backs the file up to <file>.bak first.",
    inputSchema: {
      name: z.string().min(1).describe("Header text of the column to delete."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook();
      const target = findHeader(h, args.name);
      if (!target) {
        const names = readHeaderRow(h)
          .filter((x) => x.name)
          .map((x) => x.name);
        throw new UserFacingError(
          `No column named "${args.name}". Existing columns: ${names.join(", ")}.`
        );
      }
      assertNotCanonical(target.name, "deleted");

      const backup = backupFile(h);
      deleteColumn(h, target.col);
      saveWorkbook(h);

      return textResult({
        message: `Deleted column "${target.name}".`,
        backup,
      });
    })
);
}
