/**
 * Broad-search query builder: run_job_sweep's board scan only ever covers a
 * fixed list of ~95 known companies (see sweep engine's BOARDS list). No
 * matter how often it runs, it structurally cannot surface a posting at a
 * company that isn't already on that list — the list itself is the ceiling
 * on what the board sweep can ever see.
 *
 * This tool patches that gap, but it cannot run a web search itself — this
 * server has no outbound web-search access, only the calling model does (via
 * its own WebSearch/WebFetch tools). So instead of searching, this tool
 * turns the same saved search criteria run_job_sweep already filters against
 * into a small, consistent set of ready-to-run queries, so a broad,
 * criteria-first pass happens the same way every time instead of being
 * hand-composed (or skipped) per session.
 *
 * Added 2026-09-04 after a coverage gap was flagged: the curated board list
 * can only ever find postings at companies already on it.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import {
  readSearchCriteria,
  isCriteriaComplete,
  type SearchCriteria,
} from "./searchCriteria.js";

const ATS_DOMAINS = [
  "greenhouse.io",
  "ashbyhq.com",
  "lever.co",
  "smartrecruiters.com",
  "myworkdayjobs.com",
];

function titlesPhrase(c: SearchCriteria): string {
  return c.jobTitles.map((t) => `"${t}"`).join(" OR ");
}

function locationPhrase(c: SearchCriteria): string {
  const bits: string[] = [];
  if (c.workStyle.includes("remote")) {
    bits.push(c.country ? `"remote ${c.country}"` : "remote");
  }
  if (c.workStyle.includes("hybrid") || c.workStyle.includes("onsite")) {
    if (c.city) bits.push(`"${c.city}" hybrid`);
  }
  if (bits.length) return bits.join(" OR ");
  return c.city || c.country || "";
}

export function register(server: McpServer): void {
  server.registerTool(
    "get_broad_search_queries",
    {
      title: "Get broad (non-board-list) search queries",
      description:
        "Returns a small set of ready-to-run WebSearch queries built from " +
        "the saved search criteria (see get_search_criteria), aimed at " +
        "finding postings at companies that are NOT on run_job_sweep's " +
        "fixed ~95-board list. This server has no web-search access itself " +
        "— it cannot run these queries — so the CALLING model must run " +
        "each one with its own WebSearch tool, then WebFetch any promising " +
        "result to confirm the posting is live and genuinely matches " +
        "(location, comp, domain fit) before treating it as a candidate. " +
        "Treat this as a MANDATORY step every session, not a fallback for " +
        "when run_job_sweep comes up short: run_job_sweep structurally " +
        "cannot find a company that isn't already on its list, so this is " +
        "the only mechanism that can. After verifying a hit, feed it into " +
        "discovery_sync exactly like a run_job_sweep candidate, and note " +
        "in match_assessment that it came from a broad search rather than " +
        "the board sweep. If the same company keeps turning up through " +
        "this path, it is worth adding permanently to the board sweep's " +
        "list so future runs catch it for free. Requires search criteria " +
        "to be set first, same as run_job_sweep.",
      inputSchema: {},
    },
    async () =>
      guard(() => {
        const criteria = readSearchCriteria();
        if (!isCriteriaComplete(criteria)) {
          throw new UserFacingError(
            "Search criteria are not fully set yet. Call get_search_criteria " +
              "to see what's missing, ask the user for it, save it with " +
              "update_search_criteria, then call this again."
          );
        }

        const titles = titlesPhrase(criteria);
        const loc = locationPhrase(criteria);
        const atsFilter = ATS_DOMAINS.map((d) => `site:${d}`).join(" OR ");
        const year = new Date().getFullYear();

        const queries = [
          `${titles} ${loc} ${year}`.trim(),
          `(${atsFilter}) (${titles}) ${loc}`.trim(),
        ];
        if (criteria.notes.trim()) {
          queries.push(`${titles} ${loc} ${criteria.notes.trim()}`.trim());
        }

        return textResult({
          queries,
          min_salary_hint:
            criteria.minSalary > 0
              ? `${criteria.minSalary} ${criteria.currency} — not embeddable in a search ` +
                "query; confirm from each posting instead, not from the search snippet."
              : undefined,
          instructions:
            "Run each query above with WebSearch. Skim results for companies NOT " +
            "already covered by run_job_sweep's board list — a hit at an already-" +
            "covered company adds nothing new. For anything promising, WebFetch the " +
            "actual posting (not just the search snippet) to confirm: it is currently " +
            "open, the location/work style genuinely matches, salary clears the " +
            "minimum, and the domain fit is real, not superficial. Keep this pass " +
            "lean — a handful of queries and only verifying genuinely promising hits, " +
            "not every search result. Combine surviving candidates with " +
            "run_job_sweep's before writing to discovery_sync. Finding nothing some " +
            "days is a normal, fine outcome — don't pad results to look productive.",
        });
      })
  );
}
