/**
 * Standing job-search criteria: exposes the daily-job-search-top5 SKILL.md
 * as a tool result, for clients with no Read/Bash tool access (e.g. a plain
 * MCP connection, not Claude Code) to reach the same source of truth the
 * scheduled task and /job-search-now read directly off disk.
 *
 * Reads the file fresh on every call rather than embedding a copy — SKILL.md
 * is a living document Dale edits over time (criteria, board list, gotchas),
 * and a baked-in copy here would silently drift out of sync with it.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import { SWEEP_DIR } from "./config.js";

const SKILL_FILE = path.join(SWEEP_DIR, "SKILL.md");

export function register(server: McpServer): void {
server.registerTool(
  "get_job_search_instructions",
  {
    title: "Get job-search standing instructions",
    description:
      "Return the full daily-job-search-top5 SKILL.md — the standing search " +
      "criteria (location rules, $170K+ CAD floor, role level, domain fit, " +
      "candidate-profile context used to judge match), the known-dead-end " +
      "board list, and the process this routine follows. Read this BEFORE " +
      "running a search so results match Dale's actual standing criteria " +
      "instead of improvised ones. Steps that mention sweep.py or " +
      "write_sheet.py (Bash-only) map onto run_job_sweep and discovery_sync " +
      "respectively when Bash/script access isn't available — everything " +
      "else (criteria, dedupe rules, email content rules) applies as written.",
    inputSchema: {},
  },
  async () =>
    guard(() => {
      if (!fs.existsSync(SKILL_FILE)) {
        throw new UserFacingError(`Instructions file not found at:\n  ${SKILL_FILE}`);
      }
      const text = fs.readFileSync(SKILL_FILE, "utf-8");
      return textResult(text);
    })
);
}
