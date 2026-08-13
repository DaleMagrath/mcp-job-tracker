/**
 * Discovery-sheet tools (Job_Search_Discovery.xlsx).
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import { isoToSerial, todayISO } from "./dates.js";
import { SERVER_DIR, DISCOVERY, TRACKER } from "./config.js";
import {
  openWorkbook,
  readAllRecords,
  findByCompanyPosition,
  present,
  eq,
  includesCI,
  hasColumn,
  appendRow,
  deleteRow,
  writeCell,
  backupFile,
  saveWorkbook,
  resolveFieldWrites,
} from "./workbook.js";

const execFileAsync = promisify(execFile);

/** Discovery columns set through their own parameter, not `fields`. */
const DISCOVERY_MATCH_LC = new Set<string>(["company", "position"]);
const DISCOVERY_DEFAULT_STATUS = "Open";

/**
 * Re-apply the standard formatting to the discovery workbook after an MCP write.
 * The SheetJS write handles the row data but doesn't re-serialize rich styling;
 * this openpyxl pass restores the bold header / freeze / auto-filter / links /
 * widths so MCP edits match what the daily task produces. Best-effort: returns
 * false (with a reason) if Python or openpyxl isn't available, rather than
 * failing the whole tool call — the data write already succeeded.
 */
async function reformatDiscoveryFile(): Promise<{ ok: boolean; detail?: string }> {
  const script = path.join(SERVER_DIR, "..", "scripts", "format_discovery.py");
  if (!fs.existsSync(script)) {
    return { ok: false, detail: "formatter script not found" };
  }
  for (const py of ["python", "py"]) {
    try {
      await execFileAsync(py, [script, DISCOVERY.filePath], { windowsHide: true });
      return { ok: true };
    } catch (err: any) {
      if (err?.code === "ENOENT") continue; // this interpreter isn't installed
      return { ok: false, detail: err?.stderr || err?.message || String(err) };
    }
  }
  return {
    ok: false,
    detail: "Python not found — install Python + openpyxl to keep discovery formatting.",
  };
}

export function register(server: McpServer): void {
// D1. discovery_list -------------------------------------------------
server.registerTool(
  "discovery_list",
  {
    title: "List discovery leads",
    description:
      "List job leads from the discovery sheet (Job_Search_Discovery.xlsx). " +
      "Optional filters: company (case-insensitive substring), status " +
      "(case-insensitive exact), and query (substring across every column).",
    inputSchema: {
      company: z.string().optional().describe("Company substring to filter by."),
      status: z.string().optional().describe('Exact status, e.g. "Open".'),
      query: z
        .string()
        .optional()
        .describe("Substring searched across all columns."),
    },
  },
  async (args) =>
    guard(() => {
      const h = openWorkbook(DISCOVERY);
      let records = readAllRecords(h);
      if (args.company) {
        records = records.filter((r) =>
          includesCI(r.values["Company"] ?? "", args.company!)
        );
      }
      if (args.status) {
        records = records.filter((r) => eq(r.values["Status"] ?? "", args.status!));
      }
      if (args.query) {
        const q = args.query;
        records = records.filter((r) =>
          h.order.some((c) => includesCI(r.values[c] ?? "", q))
        );
      }
      return textResult({
        file: h.spec.filePath,
        count: records.length,
        leads: records.map((r) => present(h, r)),
      });
    })
);

// D2. discovery_add --------------------------------------------------
server.registerTool(
  "discovery_add",
  {
    title: "Add discovery lead",
    description:
      "Append a new lead to the discovery sheet. Company and Position are " +
      "required; set any other columns via fields (e.g. Salary, Job Link, " +
      '"Match Assessment"). Date Found defaults to today and Status to ' +
      `"${DISCOVERY_DEFAULT_STATUS}" if not supplied. Backs the file up first.`,
    inputSchema: {
      company: z.string().min(1).describe("Company name (required)."),
      position: z.string().min(1).describe("Position / title (required)."),
      fields: z
        .record(z.string(), z.string())
        .optional()
        .describe(
          'Other discovery columns as {"Column Name": "value"}. Each name must ' +
            "be an existing discovery column."
        ),
    },
  },
  async (args) =>
    guard(async () => {
      const h = openWorkbook(DISCOVERY);
      const writes: { col: string; value: string }[] = [
        { col: "Company", value: args.company },
        { col: "Position", value: args.position },
      ];
      const fieldWrites = resolveFieldWrites(h, args.fields, DISCOVERY_MATCH_LC);
      const provided = new Set(fieldWrites.map((w) => w.col.trim().toLowerCase()));
      writes.push(...fieldWrites);

      // Sensible defaults for the two housekeeping columns, if not supplied.
      if (hasColumn(h, "Date Found") && !provided.has("date found")) {
        writes.push({ col: "Date Found", value: todayISO() });
      }
      if (hasColumn(h, "Status") && !provided.has("status")) {
        writes.push({ col: "Status", value: DISCOVERY_DEFAULT_STATUS });
      }

      const backup = backupFile(h);
      appendRow(h, writes);
      saveWorkbook(h);
      const fmt = await reformatDiscoveryFile();

      const added: Record<string, string> = {};
      for (const w of writes) if (w.value !== "") added[w.col] = w.value;
      return textResult({
        message: `Added discovery lead: ${args.company} — ${args.position}.`,
        backup,
        added,
        formatted: fmt.ok,
      });
    })
);

// D3. discovery_update -----------------------------------------------
server.registerTool(
  "discovery_update",
  {
    title: "Update discovery lead",
    description:
      "Edit a discovery lead, matched by company + position. Rename via " +
      "new_company / new_position; set any columns via fields (empty string " +
      "clears a value). Refuses missing/ambiguous matches; backs up first.",
    inputSchema: {
      company: z.string().min(1).describe("Current company (case-insensitive)."),
      position: z.string().min(1).describe("Current position (case-insensitive)."),
      new_company: z.string().optional().describe("Rename the company."),
      new_position: z.string().optional().describe("Rename the position."),
      fields: z
        .record(z.string(), z.string())
        .optional()
        .describe('Columns to set as {"Column Name": "value"}; "" clears.'),
    },
  },
  async (args) =>
    guard(async () => {
      const h = openWorkbook(DISCOVERY);
      const matches = findByCompanyPosition(
        readAllRecords(h),
        args.company,
        args.position
      );
      if (matches.length === 0) {
        throw new UserFacingError(
          `No discovery lead for "${args.company}" — "${args.position}". ` +
            `Use discovery_list to see exact values.`
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `${matches.length} discovery leads match "${args.company}" — ` +
            `"${args.position}". Refusing an ambiguous update.`
        );
      }
      const rec = matches[0];

      const updates: { col: string; value: string }[] = [];
      if (args.new_company !== undefined)
        updates.push({ col: "Company", value: args.new_company });
      if (args.new_position !== undefined)
        updates.push({ col: "Position", value: args.new_position });
      updates.push(...resolveFieldWrites(h, args.fields, DISCOVERY_MATCH_LC));

      if (updates.length === 0) {
        throw new UserFacingError(
          "No changes provided. Supply new_company, new_position, and/or fields."
        );
      }

      const changed = updates.map((u) => ({
        field: u.col,
        from: rec.values[u.col] ?? "",
        to: u.value,
      }));

      const backup = backupFile(h);
      for (const u of updates) writeCell(h, rec.row, u.col, u.value, rec.row);
      saveWorkbook(h);
      const fmt = await reformatDiscoveryFile();

      const h2 = openWorkbook(DISCOVERY);
      const after = readAllRecords(h2).find((r) => r.row === rec.row);
      return textResult({
        message: `Updated discovery lead (${changed.length} field(s)).`,
        formatted: fmt.ok,
        changed,
        lead: after ? present(h2, after) : undefined,
        backup,
      });
    })
);

// D4. discovery_delete -----------------------------------------------
server.registerTool(
  "discovery_delete",
  {
    title: "Delete discovery lead",
    description:
      "Permanently remove a discovery lead, matched by company + position, " +
      "shifting rows below it up. Refuses missing/ambiguous matches, backs up " +
      "first, and returns the deleted row so it can be re-added.",
    inputSchema: {
      company: z.string().min(1).describe("Company (case-insensitive)."),
      position: z.string().min(1).describe("Position (case-insensitive)."),
    },
  },
  async (args) =>
    guard(async () => {
      const h = openWorkbook(DISCOVERY);
      const matches = findByCompanyPosition(
        readAllRecords(h),
        args.company,
        args.position
      );
      if (matches.length === 0) {
        throw new UserFacingError(
          `No discovery lead for "${args.company}" — "${args.position}". ` +
            `Nothing deleted.`
        );
      }
      if (matches.length > 1) {
        throw new UserFacingError(
          `${matches.length} discovery leads match "${args.company}" — ` +
            `"${args.position}". Refusing an ambiguous delete.`
        );
      }
      const rec = matches[0];
      const deleted = present(h, rec);

      const backup = backupFile(h);
      deleteRow(h, rec.row);
      saveWorkbook(h);
      const fmt = await reformatDiscoveryFile();

      return textResult({
        message: `Deleted discovery lead: ${rec.values.Company} — ${rec.values.Position}.`,
        backup,
        formatted: fmt.ok,
        deletedRow: deleted,
        note: "Backed up before deleting; use discovery_add to restore.",
      });
    })
);

// D5. promote_to_tracker (discovery -> tracker) ----------------------
server.registerTool(
  "promote_to_tracker",
  {
    title: "Promote lead to tracker",
    description:
      "Turn a discovery lead into a tracked application: copies Company, " +
      "Position, Job Link, and Location into the tracker with Status " +
      '"Applied" and Date Applied = today (both overridable). Salary and Match ' +
      "Assessment are carried into Notes unless you supply your own. Optionally " +
      "removes the lead from the discovery sheet. Backs up both files it writes.",
    inputSchema: {
      company: z.string().min(1).describe("Discovery lead company (case-insensitive)."),
      position: z
        .string()
        .min(1)
        .describe("Discovery lead position (case-insensitive)."),
      resume_version: z
        .string()
        .optional()
        .describe("Resume version used for the application."),
      status: z
        .string()
        .optional()
        .describe(`Tracker status. Defaults to "Applied".`),
      date_applied: z
        .string()
        .optional()
        .describe("Date applied (YYYY-MM-DD). Defaults to today."),
      notes: z
        .string()
        .optional()
        .describe("Tracker notes. Defaults to a summary of salary + match."),
      remove_from_discovery: z
        .boolean()
        .optional()
        .describe(
          "If true, delete the lead from the discovery sheet after promoting. " +
            "Defaults to false (the lead is left in place)."
        ),
    },
  },
  async (args) =>
    guard(async () => {
      // 1) Find and validate the discovery lead first (before writing anything).
      const dh = openWorkbook(DISCOVERY);
      const dMatches = findByCompanyPosition(
        readAllRecords(dh),
        args.company,
        args.position
      );
      if (dMatches.length === 0) {
        throw new UserFacingError(
          `No discovery lead for "${args.company}" — "${args.position}". ` +
            `Use discovery_list to see exact values.`
        );
      }
      if (dMatches.length > 1) {
        throw new UserFacingError(
          `${dMatches.length} discovery leads match "${args.company}" — ` +
            `"${args.position}". Refusing an ambiguous promote.`
        );
      }
      const lead = dMatches[0];
      const company = lead.values["Company"] ?? args.company;
      const position = lead.values["Position"] ?? args.position;
      const jobLink = lead.values["Job Link"] ?? "";
      const location = lead.values["Location (Remote/Hybrid)"] ?? "";
      const salary = lead.values["Salary"] ?? "";
      const match = lead.values["Match Assessment"] ?? "";

      const dateApplied = args.date_applied ?? todayISO();
      isoToSerial(dateApplied); // validate up front
      const status = args.status ?? "Applied";
      const autoNote = [
        salary ? `Salary: ${salary}` : "",
        match ? `Match: ${match}` : "",
      ]
        .filter(Boolean)
        .join(" | ");
      const notes = args.notes ?? autoNote;

      // 2) Refuse to create a duplicate tracker row.
      const th = openWorkbook(TRACKER);
      if (findByCompanyPosition(readAllRecords(th), company, position).length) {
        throw new UserFacingError(
          `"${company}" — "${position}" is already in the tracker. ` +
            `Nothing was added. (Use update_job to change it.)`
        );
      }

      const trackerWrites: { col: string; value: string }[] = [
        { col: "Company", value: company },
        { col: "Position", value: position },
        { col: "Job Link", value: jobLink },
        { col: "Location", value: location },
        { col: "Date Applied", value: dateApplied },
        { col: "Status", value: status },
      ];
      if (args.resume_version)
        trackerWrites.push({ col: "Resume Version", value: args.resume_version });
      if (notes) trackerWrites.push({ col: "Notes", value: notes });

      const trackerBackup = backupFile(th);
      appendRow(th, trackerWrites);
      saveWorkbook(th);

      // 3) Optionally remove the lead from discovery.
      let discoveryBackup: string | null = null;
      let removed = false;
      if (args.remove_from_discovery) {
        discoveryBackup = backupFile(dh);
        deleteRow(dh, lead.row);
        saveWorkbook(dh);
        await reformatDiscoveryFile();
        removed = true;
      }

      return textResult({
        message:
          `Promoted ${company} — ${position} to the tracker ` +
          `(status "${status}", applied ${dateApplied}).`,
        trackerBackup,
        removedFromDiscovery: removed,
        discoveryBackup,
        note: removed
          ? undefined
          : "Left the discovery lead in place; pass remove_from_discovery=true " +
            "to move it, or update its status with discovery_update.",
      });
    })
);
}
