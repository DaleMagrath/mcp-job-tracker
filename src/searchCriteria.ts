/**
 * Persisted job-search criteria: location/work-style preferences, minimum
 * salary, and target job titles — the standing filters run_job_sweep judges
 * candidates against. Lives in search_criteria.json, a JOB_ROOT sibling of
 * the tracker/discovery workbooks (see resume_master.json for the same
 * pattern: a small JSON file the model reads/writes through dedicated tools
 * rather than a spreadsheet).
 *
 * First-run behavior: get_search_criteria reports which fields are still
 * unset (a fresh install has none) along with instructions telling the
 * calling model to ask the user for them and then call
 * update_search_criteria — there is no interactive prompt at the MCP layer
 * itself, so the tool result carries the prompt instead.
 */

import * as fs from "node:fs";
import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { UserFacingError, textResult, guard } from "./errors.js";
import { ensureDir, backupWithRotation } from "./workbook.js";
import { SEARCH_CRITERIA_FILE, JOB_ROOT } from "./config.js";

export const WORK_STYLES = ["remote", "hybrid", "onsite"] as const;
export type WorkStyle = (typeof WORK_STYLES)[number];

export interface SearchCriteria {
  jobTitles: string[];
  workStyle: WorkStyle[];
  city: string;
  country: string;
  minSalary: number;
  currency: string;
  notes: string;
}

const EMPTY: SearchCriteria = {
  jobTitles: [],
  workStyle: [],
  city: "",
  country: "",
  minSalary: 0,
  currency: "",
  notes: "",
};

/** Fields required before a sweep can meaningfully filter candidates. */
const REQUIRED_PROMPTS: { field: keyof SearchCriteria; prompt: string; isSet: (c: SearchCriteria) => boolean }[] = [
  {
    field: "jobTitles",
    prompt: 'Which job title(s) should the sweep search for? (e.g. "Engineering Manager", "Director of Engineering")',
    isSet: (c) => c.jobTitles.length > 0,
  },
  {
    field: "workStyle",
    prompt: "Work style — remote, hybrid, on-site, or some combination?",
    isSet: (c) => c.workStyle.length > 0,
  },
  {
    field: "city",
    prompt: "Which city (and country) should on-site/hybrid roles be near? For remote-only, the country is enough.",
    isSet: (c) => c.city.trim() !== "" || c.country.trim() !== "",
  },
  {
    field: "country",
    prompt: "Which country should the search cover?",
    isSet: (c) => c.country.trim() !== "",
  },
  {
    field: "minSalary",
    prompt: "What is the minimum acceptable annual salary (a number), and in what currency?",
    isSet: (c) => c.minSalary > 0,
  },
];

export function readSearchCriteria(): SearchCriteria {
  try {
    const raw = JSON.parse(fs.readFileSync(SEARCH_CRITERIA_FILE, "utf-8"));
    return { ...EMPTY, ...raw };
  } catch {
    return { ...EMPTY };
  }
}

export function criteriaExists(): boolean {
  return fs.existsSync(SEARCH_CRITERIA_FILE);
}

function missingFields(c: SearchCriteria) {
  return REQUIRED_PROMPTS.filter((r) => !r.isSet(c));
}

export function isCriteriaComplete(c: SearchCriteria): boolean {
  return missingFields(c).length === 0;
}

function writeSearchCriteria(c: SearchCriteria): { backup: string | null } {
  ensureDir(JOB_ROOT);
  let backup: string | null = null;
  if (fs.existsSync(SEARCH_CRITERIA_FILE)) {
    backup = backupWithRotation(SEARCH_CRITERIA_FILE);
  }
  fs.writeFileSync(SEARCH_CRITERIA_FILE, JSON.stringify(c, null, 2));
  return { backup };
}

export function register(server: McpServer): void {
  // get_search_criteria ---------------------------------------------
  server.registerTool(
    "get_search_criteria",
    {
      title: "Get saved job-search criteria",
      description:
        "Return the saved search criteria run_job_sweep filters against: " +
        "target job titles, work style (remote/hybrid/onsite), city/country, " +
        "and minimum annual salary, plus a free-text notes field for " +
        "qualitative domain-fit guidance. Call this BEFORE run_job_sweep. " +
        "If any required field is unset (a fresh install has none), the " +
        "result's `isComplete` is false and `promptsNeeded` lists exactly " +
        "what to ask the user — ask for those, then call " +
        "update_search_criteria with their answers before sweeping.",
      inputSchema: {},
    },
    async () =>
      guard(() => {
        const criteria = readSearchCriteria();
        const missing = missingFields(criteria);
        return textResult({
          isComplete: missing.length === 0,
          criteria,
          ...(missing.length
            ? {
                promptsNeeded: missing.map((m) => ({ field: m.field, ask: m.prompt })),
                instructions:
                  "Ask the user for the field(s) listed in promptsNeeded, in plain " +
                  "conversation (not necessarily one at a time), then call " +
                  "update_search_criteria with their answers. Do not guess these " +
                  "values or invent defaults.",
              }
            : {}),
        });
      })
  );

  // update_search_criteria -------------------------------------------
  server.registerTool(
    "update_search_criteria",
    {
      title: "Update job-search criteria",
      description:
        "Create or update the saved search criteria (search_criteria.json). " +
        "Only the fields you pass are changed; omit a field to leave it as " +
        "is. Use this the first time to save the user's answers to " +
        "get_search_criteria's prompts, and any time after to change a " +
        "preference. Backs up the previous file first.",
      inputSchema: {
        job_titles: z
          .array(z.string().min(1))
          .optional()
          .describe('Target job titles to search for, e.g. ["Engineering Manager", "Director of Engineering"]. Replaces the whole list.'),
        work_style: z
          .array(z.enum(WORK_STYLES))
          .optional()
          .describe('Any combination of "remote", "hybrid", "onsite". Replaces the whole list.'),
        city: z.string().optional().describe("City for on-site/hybrid roles. Pass \"\" to clear."),
        country: z.string().optional().describe("Country to search in."),
        min_salary: z
          .number()
          .nonnegative()
          .optional()
          .describe("Minimum acceptable annual salary, as a plain number (no currency symbol)."),
        currency: z.string().optional().describe('Currency the min_salary figure is in, e.g. "CAD", "USD".'),
        notes: z
          .string()
          .optional()
          .describe(
            "Free-text qualitative guidance beyond the structured fields — domain " +
              "fit, background/strengths to weigh, industries to avoid, anything " +
              "that shapes how a candidate posting should be judged. Pass \"\" to clear."
          ),
      },
    },
    async (args) =>
      guard(() => {
        if (Object.keys(args).length === 0) {
          throw new UserFacingError(
            "No changes provided. Pass at least one field to set."
          );
        }
        const current = readSearchCriteria();
        const next: SearchCriteria = {
          jobTitles: args.job_titles ?? current.jobTitles,
          workStyle: (args.work_style as WorkStyle[] | undefined) ?? current.workStyle,
          city: args.city ?? current.city,
          country: args.country ?? current.country,
          minSalary: args.min_salary ?? current.minSalary,
          currency: args.currency ?? current.currency,
          notes: args.notes ?? current.notes,
        };
        const { backup } = writeSearchCriteria(next);
        const missing = missingFields(next);
        return textResult({
          message: "Search criteria saved.",
          backup,
          isComplete: missing.length === 0,
          criteria: next,
          ...(missing.length
            ? { stillMissing: missing.map((m) => m.field) }
            : {}),
        });
      })
  );
}
