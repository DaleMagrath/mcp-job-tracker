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
 * Structure: shared modules (config / errors / dates / workbook engine) plus four
 * tool modules (trackerTools, discoveryTools, interviewPrep, documents), each of
 * which exports register(server) to attach its tools to the shared instance
 * created here.
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

const server = new McpServer(
  {
    name: "job-tracker",
    version: "1.0.0",
  },
  {
    instructions:
      "This server manages a local job search: a tracker (Job_Tracking.xlsx), a " +
      "discovery pipeline, interview prep, and resume documents (Resumes\\).\n\n" +
      "RESUMES: To create or tailor a resume, ALWAYS use the `generate_resume` " +
      "tool. It renders a correctly-formatted 2-page PDF host-side from the stable " +
      "facts in resume_master.json plus the per-posting `summary` and " +
      "`key_qualifications` you supply, saves it into Resumes\\, and returns a " +
      "`nextStep` for tracking it. Do NOT write your own python-docx/LibreOffice " +
      "script for resumes, and do NOT move documents through base64. For a file " +
      "that already exists on disk, use `save_document` with `source_path` (never " +
      "`content_base64` for anything non-trivial). After generating, offer the " +
      "returned `nextStep` (add_job / promote_to_tracker / update_job).\n\n" +
      "GMAIL: search_gmail_for_job and scan_job_updates are read-only. " +
      "draft_gmail_reply saves a Gmail draft but does not send. " +
      "send_gmail_email SENDS IMMEDIATELY and cannot be undone — only call it " +
      "with confirm: true after the user has explicitly approved the exact " +
      "to/subject/body. All four require Gmail to be authorized first (see " +
      "GMAIL_SETUP.md / npm run gmail:auth); until then they return a clear " +
      "error telling the user what to run.",
  }
);

// Each domain module registers its tools onto the shared server instance.
registerTracker(server);
registerDiscovery(server);
registerInterviewPrep(server);
registerDocuments(server);
registerGmail(server);

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
