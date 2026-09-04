/**
 * Job-board sweep tool: runs the sweep engine (sweepEngine.ts) in-process
 * and returns its new-candidate results. No subprocess, no Python — the
 * board scan, filtering, detail fetch, and dedupe all happen inside this
 * server, so the tool works identically from any MCP client.
 */

import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import { runSweep, type Candidate } from "./sweepEngine.js";
import { readSearchCriteria, isCriteriaComplete } from "./searchCriteria.js";

/** Trim a candidate to the fields worth spending context on. Drops the ATS
 *  plumbing fields (ats/token/jid) — internal to the sweep, never useful here. */
function trimCandidate(c: Candidate) {
  return {
    company: c.company,
    title: c.title,
    location: c.location,
    loc_bucket: c.locBucket,
    url: c.url,
    published: c.published,
    age_days: c.ageDays,
    stale: (c.ageDays ?? 0) >= 60,
    currency: c.currency || undefined,
    comp_low: c.compLow ?? undefined,
    comp_high: c.compHigh ?? undefined,
    clears_min_salary: c.clearsMinSalary,
    company_tracked_as: c.companyTrackedAs.length ? c.companyTrackedAs : undefined,
    // Already trimmed to ~190 chars per snippet; one is enough context.
    comp_snippet: c.compSnippets[0],
    live_status: c.liveStatus,
    live: c.live,
  };
}

export function register(server: McpServer): void {
  server.registerTool(
    "run_job_sweep",
    {
      title: "Run job board sweep",
      description:
        "Run the board sweep: probes ~95 Greenhouse/Ashby/Lever/" +
        "SmartRecruiters/Workday boards concurrently, filters to engineering-" +
        "leadership titles whose location plausibly matches the saved search " +
        "criteria, fetches salary text and posted dates, HTTP-checks posting " +
        "links, and dedupes against both spreadsheets. Takes about 60-90 " +
        "seconds. Returns only the NEW candidates (not already on the " +
        "discovery sheet or in the tracker) — still needs a judgment pass for " +
        "domain fit, comp confirmation, and hybrid-office location before any " +
        "are added via discovery_add. Salary is quoted verbatim (one snippet " +
        "per candidate); do not average a US and a local-currency band found " +
        "in the same posting — read which one applies. loc_bucket 'maybe'/" +
        "'unknown' means verify the location from the live posting before " +
        "trusting it. Requires search criteria to be set first (see " +
        "get_search_criteria) — refuses to run on an unset/incomplete " +
        "criteria file rather than filtering against guessed defaults.",
      inputSchema: {
        quick: z
          .boolean()
          .optional()
          .describe(
            "Skip boards recorded as unreachable in the last sweep's health " +
              "cache. Faster, but may miss a board that was briefly down. " +
              "Default false (probe every board)."
          ),
      },
    },
    async (args) =>
      guard(async () => {
        const criteria = readSearchCriteria();
        if (!isCriteriaComplete(criteria)) {
          throw new UserFacingError(
            "Search criteria are not fully set yet. Call get_search_criteria " +
              "to see what's missing, ask the user for it, save it with " +
              "update_search_criteria, then run this again."
          );
        }

        const result = await runSweep(criteria, { quick: args.quick ?? false });
        const fresh = result.candidates.filter((c) => c.isNew);

        return textResult({
          date: result.date,
          counts: {
            boards_ok: result.counts.boardsOk,
            boards_tried: result.counts.boardsTried,
            postings: result.counts.postings,
            leadership: result.counts.leadership,
            geo_plausible: result.counts.geoPlausible,
            new: result.counts.new,
          },
          dead_boards: result.deadBoards,
          new_candidate_count: fresh.length,
          new_candidates: fresh.map(trimCandidate),
          note:
            "Full detail for every candidate (including already-known ones this " +
            "result suppressed) is in sweep_out.json next to the tracker workbook.",
        });
      })
  );
}
