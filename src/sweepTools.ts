/**
 * Job-board sweep tool: runs sweep.py (the daily-job-search-top5 board
 * scanner) as a subprocess and returns its new-candidate results.
 *
 * Exists so the sweep can run from any MCP client (not just Claude Code's
 * Bash tool) — the subprocess lives entirely inside this server process.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import { SWEEP_SCRIPT, SWEEP_DIR, SWEEP_OUT_FILE } from "./config.js";

const execFileAsync = promisify(execFile);

/** Generous margin over sweep.py's usual 60-90s (95 boards, concurrent). */
const SWEEP_TIMEOUT_MS = 180_000;
/** sweep_out.json has run past a few hundred KB on live sweeps; keep headroom. */
const MAX_BUFFER = 20 * 1024 * 1024;

interface SweepCandidate {
  company?: string;
  title?: string;
  location?: string;
  loc_bucket?: string;
  url?: string;
  published?: string;
  age_days?: number;
  currency?: string;
  comp_low?: number | null;
  comp_high?: number | null;
  clears_170k?: boolean;
  company_tracked_as?: string[];
  comp_snippets?: string[];
  live_status?: number | null;
  live?: string;
  is_new?: boolean;
}

/** Trim a candidate to the fields worth spending context on. Drops the ATS
 *  plumbing fields (ats/token/jid) — internal to sweep.py, never useful here. */
function trimCandidate(c: SweepCandidate) {
  return {
    company: c.company,
    title: c.title,
    location: c.location,
    loc_bucket: c.loc_bucket,
    url: c.url,
    published: c.published,
    age_days: c.age_days,
    stale: (c.age_days ?? 0) >= 60,
    currency: c.currency || undefined,
    comp_low: c.comp_low ?? undefined,
    comp_high: c.comp_high ?? undefined,
    clears_170k: c.clears_170k,
    company_tracked_as: c.company_tracked_as?.length ? c.company_tracked_as : undefined,
    // Sweep.py already trims each snippet to ~250 chars; one is enough context.
    comp_snippet: c.comp_snippets?.[0],
    live_status: c.live_status,
    live: c.live,
  };
}

export function register(server: McpServer): void {
server.registerTool(
  "run_job_sweep",
  {
    title: "Run job board sweep",
    description:
      "Run the daily-job-search-top5 board sweep (sweep.py): probes ~95 " +
      "Greenhouse/Ashby/Lever/SmartRecruiters/Workday boards concurrently, " +
      "filters to senior engineering-leadership titles that are Canada-" +
      "eligible, fetches salary text and posted dates, HTTP-checks posting " +
      "links, and dedupes against both spreadsheets. Takes about 60-90 " +
      "seconds. Returns only the NEW candidates (not already on the " +
      "discovery sheet or in the tracker) — still needs a judgment pass for " +
      "domain fit, comp confirmation, and hybrid-office location before any " +
      "are added via discovery_add. Salary is quoted verbatim (one snippet " +
      "per candidate); do not average a US and a CAD band found in the same " +
      "posting — read which one applies. loc_bucket 'maybe'/'unknown' means " +
      "verify the location from the live posting before trusting it.",
    inputSchema: {},
  },
  async () =>
    guard(async () => {
      if (!fs.existsSync(SWEEP_SCRIPT)) {
        throw new UserFacingError(`sweep.py not found at:\n  ${SWEEP_SCRIPT}`);
      }

      let stderr = "";
      let ranWith: string | null = null;
      for (const py of ["python", "py"]) {
        try {
          const result = await execFileAsync(py, [SWEEP_SCRIPT], {
            cwd: SWEEP_DIR,
            windowsHide: true,
            timeout: SWEEP_TIMEOUT_MS,
            maxBuffer: MAX_BUFFER,
            env: {
              ...process.env,
              // sweep.py's own print()s crash on Windows consoles whose active
              // code page can't encode a scraped location string (seen live,
              // 2026-08-20: cp1252 vs a Czech city name). None of that printed
              // text is used below anyway — sweep_out.json is the real output —
              // but forcing UTF-8 keeps the subprocess from dying on stdout.
              PYTHONIOENCODING: "utf-8",
            },
          });
          stderr = result.stderr || "";
          ranWith = py;
          break;
        } catch (err: any) {
          if (err?.code === "ENOENT") continue; // this interpreter isn't installed
          if (err?.killed || err?.signal === "SIGTERM") {
            throw new UserFacingError(
              `sweep.py did not finish within ${SWEEP_TIMEOUT_MS / 1000}s and was ` +
                `killed. It usually takes 60-90s; a run this slow suggests a board ` +
                `is hanging. Try again, or check board_health.json for the slow one.`
            );
          }
          // A non-zero exit is still worth surfacing sweep_out.json for — the
          // 2026-08-20 crash happened AFTER the file was written, mid-printout.
          if (fs.existsSync(SWEEP_OUT_FILE)) {
            stderr = err?.stderr || err?.message || String(err);
            ranWith = py;
            break;
          }
          throw new UserFacingError(
            `sweep.py failed before writing any output: ${err?.stderr || err?.message || String(err)}`
          );
        }
      }
      if (!ranWith) {
        throw new UserFacingError(
          "Neither `python` nor `py` was found on PATH — sweep.py needs a Python " +
            "interpreter installed to run."
        );
      }

      if (!fs.existsSync(SWEEP_OUT_FILE)) {
        throw new UserFacingError(
          `sweep.py ran but produced no output at:\n  ${SWEEP_OUT_FILE}` +
            (stderr ? `\n\nstderr:\n${stderr.slice(0, 2000)}` : "")
        );
      }
      const raw = JSON.parse(fs.readFileSync(SWEEP_OUT_FILE, "utf-8"));
      const all: SweepCandidate[] = raw.candidates || [];
      const fresh = all.filter((c) => c.is_new);

      return textResult({
        date: raw.date,
        counts: raw.counts,
        dead_boards: raw.dead_boards,
        new_candidate_count: fresh.length,
        new_candidates: fresh.map(trimCandidate),
        note:
          "Full detail for every candidate (including already-known ones sweep " +
          `suppressed) is in ${path.basename(SWEEP_OUT_FILE)} next to sweep.py.`,
      });
    })
);
}
