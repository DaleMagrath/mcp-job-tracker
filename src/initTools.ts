/**
 * init_job_tracker_files: one-time bootstrap for a fresh machine/profile.
 *
 * Creates the Job Tracking root folder if it's missing, and creates
 * Job_Tracking.xlsx and/or Job_Search_Discovery.xlsx from a blank template
 * (header row only, the same columns the rest of the server expects) for
 * whichever of the two doesn't already exist. Never touches a file that's
 * already there — this is a bootstrap, not a reset.
 */

import * as fs from "node:fs";
import * as XLSX from "xlsx";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { textResult, guard } from "./errors.js";
import { ensureDir } from "./workbook.js";
import { JOB_ROOT, TRACKER, DISCOVERY, COLUMNS, DISCOVERY_COLUMNS } from "./config.js";
import { formatWorkbookFile } from "./xlsxFormat.js";

/** Write a blank workbook with just a header row for the given columns. */
function createTemplateWorkbook(filePath: string, columns: readonly string[]): void {
  const ws = XLSX.utils.aoa_to_sheet([[...columns]]);
  ws["!cols"] = columns.map(() => ({ wch: 22 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  // Same write options as saveWorkbook(): no cellStyles (see workbook.ts notes).
  XLSX.writeFile(wb, filePath, {});
}

export function register(server: McpServer): void {
  server.registerTool(
    "init_job_tracker_files",
    {
      title: "Initialize job tracker folder + template spreadsheets",
      description:
        "Bootstrap a fresh machine/profile: creates the Job Tracking root " +
        "folder if it doesn't exist, and creates Job_Tracking.xlsx and/or " +
        "Job_Search_Discovery.xlsx from a blank template (header row only, " +
        "with the standard bold-header/frozen-row/auto-filter formatting) " +
        "for whichever of the two is missing. Never overwrites or modifies a " +
        "file that already exists — this only fills in what's absent, so " +
        "it's safe to call any time and it just reports what it found. Run " +
        "this before any other job-tracker tool on a machine where the " +
        "Job Tracking folder hasn't been set up yet.",
      inputSchema: {},
    },
    async () =>
      guard(async () => {
        const folderExisted = fs.existsSync(JOB_ROOT);
        ensureDir(JOB_ROOT);

        const results: Record<string, string> = {};

        for (const [label, spec, columns] of [
          ["tracker", TRACKER, COLUMNS],
          ["discovery", DISCOVERY, DISCOVERY_COLUMNS],
        ] as const) {
          if (fs.existsSync(spec.filePath)) {
            results[label] = `already existed — left untouched: ${spec.filePath}`;
            continue;
          }
          createTemplateWorkbook(spec.filePath, columns);
          const fmt = await formatWorkbookFile(spec.filePath);
          results[label] = fmt.ok
            ? `created: ${spec.filePath}`
            : `created (plain formatting — ${fmt.detail}): ${spec.filePath}`;
        }

        return textResult({
          folder: JOB_ROOT,
          folderCreated: !folderExisted,
          tracker: results.tracker,
          discovery: results.discovery,
        });
      })
  );
}
