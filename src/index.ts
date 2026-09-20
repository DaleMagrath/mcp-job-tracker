#!/usr/bin/env node
/**
 * Job Search Assistant — MCP server
 *
 * A stdio MCP server that lets Claude read and update a job-search tracker
 * (Job_Tracking.xlsx), a discovery sheet, an interview-prep note, and saved
 * resume documents through natural conversation.
 *
 * Transport   : stdio (local process; no networking/auth).
 * Config      : file paths from env vars (JOB_TRACKER_FILE, JOB_DISCOVERY_FILE,
 *               JOB_INTERVIEW_PREP_FILE, JOB_RESUMES_DIR), else tracker siblings.
 *
 * Design notes
 *  - Workbooks are re-read fresh on every tool call so external edits (e.g. the
 *    user editing in Excel) are always reflected. There is no in-memory cache.
 *  - Dates are stored as Excel date serials formatted "d-mmm-yy"; we read serials
 *    directly (TZ-free integer math) and write new dates with the same format.
 *  - Every write first makes a rotating timestamped backup (in a .backups folder).
 *
 * Structure: shared modules (config / errors / dates / workbook engine /
 * matching) plus one file per domain (trackerTools, discoveryTools,
 * interviewPrep, documents, gmailTools, sweepTools + sweepEngine,
 * discoverySync, searchCriteria, initTools), each exporting register(server)
 * to attach its tools to the shared instance created here. No Python or
 * other external interpreter is required by anything in this list.
 */

import * as fs from "node:fs";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  TRACKER,
  DISCOVERY,
  INTERVIEW_PREP_FILE,
  RESUMES_DIR,
} from "./config.js";
import { register as registerDocuments } from "./documents.js";
import { register as registerInterviewPrep } from "./interviewPrep.js";
import { register as registerTracker } from "./trackerTools.js";
import { register as registerDiscovery } from "./discoveryTools.js";
import { register as registerGmail } from "./gmailTools.js";
import { register as registerSweep } from "./sweepTools.js";
import { register as registerBroadSearch } from "./broadSearchTool.js";
import { register as registerDiscoverySync } from "./discoverySync.js";
import { register as registerSearchCriteria } from "./searchCriteria.js";
import { register as registerInit } from "./initTools.js";
import { register as registerSetup } from "./setupTools.js";

const server = new McpServer(
  {
    name: "job-tracker",
    version: "1.0.0",
  },
  {
    instructions:
      "This server manages a local job search: a tracker (Job_Tracking.xlsx), a " +
      "discovery pipeline, interview prep, and resume documents (Resumes\\).\n\n" +
      "SETUP: on an unfamiliar or freshly-installed machine, call check_setup " +
      "first — it's read-only and reports in one call which files/folders " +
      "exist, whether search criteria and the resume master are set up, " +
      "whether Gmail is authorized, and whether a PDF converter is available, " +
      "with a `nextSteps` list telling you exactly what to do about anything " +
      "missing. Cheaper than discovering the same gaps one tool call at a " +
      "time.\n\n" +
      "RESUMES: To create or tailor a resume, ALWAYS use the `generate_resume` " +
      "tool. It renders a correctly-formatted 2-page resume host-side (PDF when " +
      "LibreOffice or Word is installed, else .docx with a clear note) from the " +
      "stable facts in resume_master.json plus the per-posting `summary` and " +
      "`key_qualifications` you supply, saves it into Resumes\\, and returns a " +
      "`nextStep` for tracking it. Do NOT write your own resume-generation " +
      "script, and do NOT move documents through base64. For a file " +
      "that already exists on disk, use `save_document` with `source_path` (never " +
      "`content_base64` for anything non-trivial). After generating, offer the " +
      "returned `nextStep` (add_job / promote_to_tracker / update_job).\n\n" +
      "NO RESUME MASTER YET: generate_resume / get_resume_master_field / " +
      "update_resume_master_field / list_resume_master_structure all error " +
      "clearly if resume_master.json doesn't exist. On that error (or on a " +
      "fresh install), ask the user to upload their master resume — the one " +
      "they'd customize per application — as a PDF or DOCX. Save it (e.g. " +
      "`save_document` with `source_path`), extract its text (`read_document`), " +
      "structure the REAL content into resume_master.json's fields (name, " +
      "contact, default_summary, default_key_qualifications, experience[], " +
      "education[], skills[], optionally projects[]/certifications[]) — never " +
      "fabricate facts, ask the user to fill in anything genuinely missing — " +
      "and save it with `create_resume_master`. That tool refuses to run if " +
      "resume_master.json already exists, so it only ever fires once.\n\n" +
      "GMAIL: search_gmail_for_job and scan_job_updates are read-only. " +
      "draft_gmail_reply saves a Gmail draft but does not send. " +
      "send_gmail_email SENDS IMMEDIATELY and cannot be undone — only call it " +
      "with confirm: true after the user has explicitly approved the exact " +
      "to/subject/body. All four require Gmail to be authorized first (see " +
      "GMAIL_SETUP.md / npm run gmail:auth); until then they return a clear " +
      "error telling the user what to run.\n\n" +
      "SWEEP: run_job_sweep runs the board scanner in-process (no Python/" +
      "subprocess dependency) and returns only new (not already known) " +
      "candidates, ~60-90s. It filters mechanically against the saved search " +
      "criteria — YOUR judgment is still required for domain fit, comp " +
      "confirmation, and hybrid-office location before calling discovery_add. " +
      "IMPORTANT LIMITATION: it only probes a fixed list of ~95 known " +
      "companies' ATS boards — it structurally cannot find a posting at a " +
      "company that isn't already on that list, no matter how often it " +
      "runs. That is what get_broad_search_queries is for (see below).\n\n" +
      "BROAD SEARCH: get_broad_search_queries returns ready-to-run WebSearch " +
      "queries built from the same saved search criteria, aimed at " +
      "companies outside run_job_sweep's board list. This server has no " +
      "web-search access itself, so the calling model must run the queries " +
      "with its own WebSearch tool and WebFetch anything promising to " +
      "verify it before treating it as a candidate. This is a MANDATORY " +
      "part of the daily routine, not a fallback for when run_job_sweep " +
      "comes up short — skipping it means silently limiting the search to " +
      "the same ~95 companies forever.\n\n" +
      "DAILY ROUTINE: get_search_criteria + run_job_sweep + " +
      "get_broad_search_queries + discovery_sync together are the whole " +
      "routine, on any client (no Bash/script access needed, no scheduler " +
      "needed — just call them in order). Call get_search_criteria first — " +
      "it holds the standing search criteria (job titles, work style, " +
      "city/country, minimum salary, plus free-text notes) that decide " +
      "which candidates actually qualify, from either source. On a fresh " +
      "install nothing is saved yet: get_search_criteria's result says so " +
      "and lists exactly what to ask the user — collect those answers and " +
      "call update_search_criteria before sweeping, rather than guessing or " +
      "inventing defaults. Then run_job_sweep to find candidates from the " +
      "known-board list, and get_broad_search_queries plus your own " +
      "WebSearch/WebFetch calls to find candidates beyond it. Judge all of " +
      "them against the saved criteria, then call discovery_sync once with " +
      "the ones that qualify (or an empty array to run housekeeping alone) " +
      "— it drops acted rows, flags stale ones, and appends the rest with " +
      "the same dedupe rules, in one call. Prefer it over repeated " +
      "discovery_add calls when processing results, since only " +
      "discovery_sync also does the housekeeping pass.",
  }
);

// Each domain module registers its tools onto the shared server instance.
registerTracker(server);
registerDiscovery(server);
registerInterviewPrep(server);
registerDocuments(server);
registerGmail(server);
registerSweep(server);
registerBroadSearch(server);
registerDiscoverySync(server);
registerSearchCriteria(server);
registerInit(server);
registerSetup(server);

/* ------------------------------------------------------------------ */
/* Boot                                                               */
/* ------------------------------------------------------------------ */

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // stderr is safe for logging on stdio transport (stdout carries the protocol).
  const status = (p: string) => (fs.existsSync(p) ? "" : "  (NOT FOUND yet)");
  console.error(
    `job-tracker MCP server ready.\n` +
      `  tracker:   ${TRACKER.filePath}${status(TRACKER.filePath)}\n` +
      `  discovery: ${DISCOVERY.filePath}${status(DISCOVERY.filePath)}\n` +
      `  prep:      ${INTERVIEW_PREP_FILE}${status(INTERVIEW_PREP_FILE)}\n` +
      `  resumes:   ${RESUMES_DIR}${status(RESUMES_DIR)}`
  );
}

main().catch((err) => {
  console.error("Fatal error starting job-tracker MCP server:", err);
  process.exit(1);
});
