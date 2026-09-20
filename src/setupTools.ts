/**
 * check_setup: one-call readiness report for a machine/profile.
 *
 * Every piece of state this reports on is already checked, separately, by
 * some other tool the first time it's used (missing resume master, missing
 * search criteria, missing Gmail token, missing PDF converter, missing
 * workbook files) — this just surfaces all of them together up front, so a
 * fresh install (or a "why isn't X working" question) doesn't require
 * triggering five different tools one at a time to find out what's missing.
 * Read-only; changes nothing.
 */

import * as fs from "node:fs";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { textResult, guard } from "./errors.js";
import {
  JOB_ROOT,
  TRACKER,
  DISCOVERY,
  INTERVIEW_PREP_FILE,
  RESUMES_DIR,
  RESUME_MASTER_FILE,
  SEARCH_CRITERIA_FILE,
} from "./config.js";
import { GMAIL_CREDENTIALS_FILE, GMAIL_TOKEN_FILE } from "./gmailConfig.js";
import { readSearchCriteria, criteriaExists, isCriteriaComplete } from "./searchCriteria.js";
import { detectPdfConverter } from "./documents.js";

export function register(server: McpServer): void {
  server.registerTool(
    "check_setup",
    {
      title: "Check job-tracker setup",
      description:
        "Read-only readiness report for this machine/profile: which files " +
        "exist (tracker, discovery, interview prep, Resumes folder, resume " +
        "master, search criteria), whether search criteria are complete " +
        "enough to run the daily routine, whether Gmail is authorized, and " +
        "whether a PDF converter (LibreOffice or Word) is available for " +
        "generate_resume's default PDF output. Changes nothing. Run this " +
        "first on an unfamiliar or freshly-set-up machine instead of " +
        "discovering gaps one tool call at a time — `nextSteps` lists " +
        "exactly what to do about anything missing.",
      inputSchema: {},
    },
    async () =>
      guard(async () => {
        const folderExists = fs.existsSync(JOB_ROOT);
        const trackerExists = fs.existsSync(TRACKER.filePath);
        const discoveryExists = fs.existsSync(DISCOVERY.filePath);
        const interviewPrepExists = fs.existsSync(INTERVIEW_PREP_FILE);
        const resumesDirExists = fs.existsSync(RESUMES_DIR);
        const resumeMasterExists = fs.existsSync(RESUME_MASTER_FILE);

        const searchCriteriaFileExists = criteriaExists();
        const criteria = readSearchCriteria();
        const searchCriteriaComplete = searchCriteriaFileExists && isCriteriaComplete(criteria);

        const gmailCredsExist = fs.existsSync(GMAIL_CREDENTIALS_FILE);
        const gmailTokenExists = fs.existsSync(GMAIL_TOKEN_FILE);
        const gmailAuthorized = gmailCredsExist && gmailTokenExists;

        const pdfConverter = await detectPdfConverter();

        const nextSteps: string[] = [];
        if (!folderExists || !trackerExists || !discoveryExists) {
          nextSteps.push(
            "Call init_job_tracker_files to create the Job Tracking folder and " +
              "template spreadsheets."
          );
        }
        if (!resumeMasterExists) {
          nextSteps.push(
            "No resume master yet — ask the user to upload their master resume " +
              "(PDF/DOCX), extract it, and save it with create_resume_master " +
              "before generate_resume can be used."
          );
        }
        if (!searchCriteriaComplete) {
          nextSteps.push(
            "Search criteria are unset or incomplete — call get_search_criteria " +
              "to see what's missing, then update_search_criteria, before " +
              "running the daily sweep routine."
          );
        }
        if (!pdfConverter.available) {
          nextSteps.push(
            "No PDF converter (LibreOffice or Microsoft Word) was found — " +
              "generate_resume will fall back to .docx output until one is " +
              "installed. This is optional; .docx works fine on its own."
          );
        }
        if (!gmailAuthorized) {
          nextSteps.push(
            "Gmail isn't authorized — optional. Run \"npm run gmail:auth\" from " +
              "the project folder if Gmail search/draft/send features are wanted."
          );
        }

        return textResult({
          jobRoot: JOB_ROOT,
          folderExists,
          tracker: { path: TRACKER.filePath, exists: trackerExists },
          discovery: { path: DISCOVERY.filePath, exists: discoveryExists },
          interviewPrep: { path: INTERVIEW_PREP_FILE, exists: interviewPrepExists },
          resumesDir: { path: RESUMES_DIR, exists: resumesDirExists },
          resumeMaster: { path: RESUME_MASTER_FILE, exists: resumeMasterExists },
          searchCriteria: {
            path: SEARCH_CRITERIA_FILE,
            exists: searchCriteriaFileExists,
            isComplete: searchCriteriaComplete,
          },
          gmail: {
            credentialsFile: { path: GMAIL_CREDENTIALS_FILE, exists: gmailCredsExist },
            tokenFile: { path: GMAIL_TOKEN_FILE, exists: gmailTokenExists },
            authorized: gmailAuthorized,
          },
          pdfConverter,
          ready: nextSteps.length === 0,
          nextSteps,
        });
      })
  );
}
