/**
 * Configuration & schema: column definitions, sheet specs, path resolvers.
 */

import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

/** Directory of the running server file (dist/), for locating sibling scripts.
 *  Computed here so it compiles to the dist root (dist/config.js), keeping the
 *  "../scripts/format_discovery.py" lookup correct. */
export const SERVER_DIR = path.dirname(fileURLToPath(import.meta.url));

/**
 * The base columns of a standard tracker, in their default order. These are
 * always required to exist and are protected from rename/delete, because the
 * tools reference them by name (matching, status, dates). The schema is
 * otherwise dynamic: any *extra* columns present in the sheet are picked up
 * automatically — shown in listings and settable per job via `extra_fields`.
 */
export const COLUMNS = [
  "Company",
  "Position",
  "Job Link",
  "Location",
  "Resume Version",
  "Contact/Referral",
  "Date Applied",
  "Status",
  "Next Follow-Up",
  "Notes",
] as const;
export type Column = (typeof COLUMNS)[number];

/**
 * The discovery / top-of-funnel sheet's columns, in order. Leads live here
 * before you apply; a promoted lead becomes a row in the tracker. Stored as a
 * separate workbook with its own schema.
 */
export const DISCOVERY_COLUMNS = [
  "Date Found",
  "Company",
  "Position",
  "Location (Remote/Hybrid)",
  "Salary",
  "Salary Confidence (Confirmed/Estimated)",
  "Job Link",
  "Source (Company Careers Page/LinkedIn/Aggregator)",
  "Posted Date / Days Since Posted",
  "Known Gap Flag",
  "Match Assessment",
  "Status",
] as const;

/**
 * Describes one workbook the server manages: where it lives, its required/
 * protected base columns, which columns hold Excel date serials, and which
 * columns identify a row. The read/write engine is otherwise schema-agnostic,
 * so a new sheet is just a new spec.
 */
export interface SheetSpec {
  /** short human label used in messages, e.g. "tracker" / "discovery". */
  label: string;
  filePath: string;
  /** required to exist; for the tracker, also protected from rename/delete. */
  baseColumns: readonly string[];
  /** header names (lowercased) whose cells are Excel date serials. */
  dateColumnsLC: ReadonlySet<string>;
  /** columns used to identify a single row (e.g. Company + Position). */
  matchColumns: readonly string[];
}

export function isDateColumn(spec: SheetSpec, name: string): boolean {
  return spec.dateColumnsLC.has(name.trim().toLowerCase());
}

export const DEFAULT_STATUS = "Awaiting Response";

/** Statuses that get suggested to the model; the field is free text, though. */
export const KNOWN_STATUSES = [
  "Applied",
  "Awaiting Response",
  "Informal - Referral Sent",
  "Interviewing",
  "Offer",
  "Declined",
  "Withdrawn",
  "Closed - No Longer Available",
];

function resolveTrackerPath(): string {
  const fromEnv = process.env.JOB_TRACKER_FILE;
  const fromArg = process.argv[2];
  const chosen =
    (fromEnv && fromEnv.trim()) ||
    (fromArg && fromArg.trim()) ||
    path.join(os.homedir(), "Documents", "Job_Tracking.xlsx");
  return path.resolve(chosen);
}

/** Discovery file: JOB_DISCOVERY_FILE, else a sibling of the tracker file. */
function resolveDiscoveryPath(trackerPath: string): string {
  const fromEnv = process.env.JOB_DISCOVERY_FILE;
  const chosen =
    (fromEnv && fromEnv.trim()) ||
    path.join(path.dirname(trackerPath), "Job_Search_Discovery.xlsx");
  return path.resolve(chosen);
}

/** Interview-prep markdown: JOB_INTERVIEW_PREP_FILE, else a tracker sibling. */
function resolveInterviewPrepPath(trackerPath: string): string {
  const fromEnv = process.env.JOB_INTERVIEW_PREP_FILE;
  const chosen =
    (fromEnv && fromEnv.trim()) ||
    path.join(path.dirname(trackerPath), "Interview_Prep_QA.md");
  return path.resolve(chosen);
}

/** Saved-documents folder: JOB_RESUMES_DIR, else a "Resumes" tracker sibling. */
function resolveResumesDir(trackerPath: string): string {
  const fromEnv = process.env.JOB_RESUMES_DIR;
  const chosen =
    (fromEnv && fromEnv.trim()) ||
    path.join(path.dirname(trackerPath), "Resumes");
  return path.resolve(chosen);
}

/** Resume master data (stable facts): JOB_RESUME_MASTER_FILE, else a sibling. */
function resolveResumeMasterPath(trackerPath: string): string {
  const fromEnv = process.env.JOB_RESUME_MASTER_FILE;
  const chosen =
    (fromEnv && fromEnv.trim()) ||
    path.join(path.dirname(trackerPath), "resume_master.json");
  return path.resolve(chosen);
}

export const TRACKER_PATH = resolveTrackerPath();

/** The Job Tracking root folder (the tracker file's directory). read_text_file
 *  is scoped to this tree; nothing may be read outside it. */
export const JOB_ROOT = path.dirname(TRACKER_PATH);

/** JSON holding the stable resume facts that generate_resume tailors from. */
export const RESUME_MASTER_FILE = resolveResumeMasterPath(TRACKER_PATH);

/** The interview-prep Q&A markdown file (plain text, not a workbook). */
export const INTERVIEW_PREP_FILE = resolveInterviewPrepPath(TRACKER_PATH);
export const INTERVIEW_PREP_TITLE = "# Interview Prep — Q&A Reference";

/** Folder where generated resumes / cover letters are saved and printed from. */
export const RESUMES_DIR = resolveResumesDir(TRACKER_PATH);

/** The tracker workbook (Job_Tracking.xlsx). */
export const TRACKER: SheetSpec = {
  label: "tracker",
  filePath: TRACKER_PATH,
  baseColumns: COLUMNS,
  dateColumnsLC: new Set(["date applied", "next follow-up"]),
  matchColumns: ["Company", "Position"],
};

/** The discovery workbook (Job_Search_Discovery.xlsx); dates stored as text. */
export const DISCOVERY: SheetSpec = {
  label: "discovery",
  filePath: resolveDiscoveryPath(TRACKER_PATH),
  baseColumns: DISCOVERY_COLUMNS,
  dateColumnsLC: new Set<string>(),
  matchColumns: ["Company", "Position"],
};
