/**
 * Job-board sweep engine — ported from the former external sweep.py (Python)
 * so run_job_sweep has no subprocess/interpreter dependency and works on any
 * machine this server is installed on.
 *
 * One call does every mechanical step: probe all boards concurrently, filter
 * to engineering-leadership titles whose location plausibly matches the
 * saved search criteria, fetch salary/posted-date detail for survivors,
 * dedupe against both spreadsheets, HTTP-check surviving links, and return a
 * ranked shortlist. Collapsing this into one in-process call (rather than
 * ~40 separate board-by-board tool calls) is the whole point — each
 * round-trip is an opportunity for a slow board to stall the caller.
 *
 * Reads (never writes) the tracker/discovery workbooks.
 */

import * as fs from "node:fs";
import { TRACKER, DISCOVERY } from "./config.js";
import { openWorkbook, readAllRecords } from "./workbook.js";
import { norm, normUrl, linksIn, keyOf } from "./matching.js";
import { SWEEP_HEALTH_FILE, SWEEP_OUT_FILE } from "./config.js";
import type { SearchCriteria } from "./searchCriteria.js";

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
  "(KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36";

/* ------------------------------------------------------------------ */
/* Board registry                                                     */
/* ------------------------------------------------------------------ */
// (display name, ats, token). Priority vendors first — workflow/PPM/work
// management, then dev tools / enterprise SaaS, then large-remote-employer
// and Canadian-SaaS boards seen productive in past runs. Tokens that never
// return data get recorded in board_health.json and skipped by quick=true
// on later runs. Add a company here, not in a prompt.
export type Ats = "greenhouse" | "ashby" | "lever" | "smartrecruiters" | "workday";
export type BoardEntry = readonly [name: string, ats: Ats, token: string];

export const BOARDS: BoardEntry[] = [
  // --- priority: workflow / PPM / work management / integration ---
  ["Wrike", "greenhouse", "wrike"],
  ["Asana", "greenhouse", "asana"],
  ["Smartsheet", "greenhouse", "smartsheet"],
  ["Airtable", "greenhouse", "airtable"],
  ["ClickUp", "ashby", "clickup"],
  ["Loopio", "ashby", "loopio"],
  ["Camunda", "ashby", "camunda"],
  ["Temporal", "ashby", "temporal"],
  ["Workato", "greenhouse", "workato"],
  ["Nintex", "greenhouse", "nintex"],
  ["Appian", "greenhouse", "appian"],
  ["Celonis", "greenhouse", "celonis"],
  ["Zapier", "ashby", "zapier"],
  ["Celigo", "greenhouse", "celigo"],
  ["n8n", "ashby", "n8n"],
  ["Merge", "ashby", "merge"],
  ["Prefect", "ashby", "prefect"],
  ["Airbyte", "ashby", "airbyte"],
  ["Astronomer", "ashby", "astronomer"],
  ["Fivetran", "greenhouse", "fivetran"],
  // --- priority: dev tools / enterprise SaaS ---
  ["GitLab", "greenhouse", "gitlab"],
  ["1Password", "ashby", "1password"],
  ["Docker", "ashby", "docker"],
  ["PagerDuty", "greenhouse", "pagerduty"],
  ["Vanta", "ashby", "vanta"],
  ["Render", "ashby", "render"],
  ["Supabase", "ashby", "supabase"],
  ["Sourcegraph", "greenhouse", "sourcegraph91"],
  ["LaunchDarkly", "greenhouse", "launchdarkly"],
  ["CircleCI", "greenhouse", "circleci"],
  ["Grafana Labs", "greenhouse", "grafanalabs"],
  ["Sentry", "ashby", "sentry"],
  ["Postman", "greenhouse", "postman"],
  ["Confluent", "ashby", "confluent"],
  ["MongoDB", "greenhouse", "mongodb"],
  ["Elastic", "greenhouse", "elastic"],
  ["Cloudflare", "greenhouse", "cloudflare"],
  ["Vercel", "greenhouse", "vercel"],
  ["Netlify", "greenhouse", "netlify"],
  ["Linear", "ashby", "linear"],
  ["Datadog", "greenhouse", "datadog"],
  ["New Relic", "greenhouse", "newrelic"],
  ["JFrog", "greenhouse", "jfrog"],
  ["Cockroach Labs", "greenhouse", "cockroachlabs"],
  ["Redis", "ashby", "redis"],
  ["Neo4j", "greenhouse", "neo4j"],
  ["Chainguard", "greenhouse", "chainguard"],
  ["Snowflake", "ashby", "snowflake"],
  ["Okta", "greenhouse", "okta"],
  ["Box", "greenhouse", "boxinc"],
  ["Amplitude", "ashby", "amplitude"],
  ["Mixpanel", "greenhouse", "mixpanel"],
  ["Webflow", "greenhouse", "webflow"],
  ["Calendly", "greenhouse", "calendly"],
  // --- Canadian SaaS / previously productive ---
  ["Jane", "ashby", "jane"],
  ["Jobber", "ashby", "jobber"],
  ["Veeva", "lever", "veeva"],
  ["Cohere", "ashby", "cohere"],
  ["D2L", "greenhouse", "d2l"],
  ["Docebo", "ashby", "docebo"],
  ["Geotab", "greenhouse", "geotab"],
  ["Hootsuite", "greenhouse", "hootsuite"],
  ["Later", "greenhouse", "later"],
  ["Tulip", "greenhouse", "tulip"],
  ["Klue", "ashby", "klue"],
  ["Assent", "smartrecruiters", "Assent"],
  ["PointClickCare", "lever", "pointclickcare"],
  ["Wealthsimple", "ashby", "wealthsimple"],
  ["Achievers", "lever", "achievers"],
  ["Benevity", "ashby", "benevity"],
  ["Thinkific", "ashby", "thinkific"],
  ["Xero", "ashby", "xero"],
  ["Wattpad", "lever", "wattpad"],
  ["Nylas", "ashby", "nylas"],
  // Ashby token 'float' is Float FINANCIAL — a Toronto fintech (corporate
  // cards, expense mgmt). NOT float.com, the resource-management/scheduling
  // vendor; that mislabel once scored a fintech board as a top PPM fit.
  // float.com is not on Ashby under any token and isn't swept.
  ["Float Financial", "ashby", "float"],
  // --- Workday-hosted boards ---
  // Token is 'tenant:host:site', NOT a bare slug — the site id is an
  // arbitrary per-company string that can't be guessed. See fetchWorkday().
  ["Clio", "workday", "clio:wd3:ClioCareerSite"],
  ["Zendesk", "workday", "zendesk:wd1:zendesk"],
  // --- large remote employers seen productive in past runs ---
  ["Instacart", "greenhouse", "instacart"],
  ["Twilio", "greenhouse", "twilio"],
  ["Stripe", "greenhouse", "stripe"],
  ["Affirm", "greenhouse", "affirm"],
  ["Lyft", "greenhouse", "lyft"],
  ["Databricks", "greenhouse", "databricks"],
  ["Dropbox", "greenhouse", "dropbox"],
  ["Notion", "ashby", "notion"],
  ["Figma", "greenhouse", "figma"],
  ["Miro", "ashby", "miro"],
  ["Samsara", "greenhouse", "samsara"],
  ["HubSpot", "greenhouse", "hubspotjobs"],
  ["Reddit", "greenhouse", "reddit"],
  ["Discord", "greenhouse", "discord"],
  ["Pinterest", "greenhouse", "pinterest"],
  ["Faire", "greenhouse", "faire"],
  ["Remote.com", "greenhouse", "remotecom"],
  ["Oyster", "ashby", "oyster"],
];

// Boards deliberately NOT in the list, so nobody re-adds them on a hunch:
//   Pega, monday.com, Boomi, Kinaxis — self-hosted careers sites, no public
//     ATS JSON (verified against ~20 ATS fingerprints). Pega alone is worth a
//     manual look when a sweep comes up short: ~107 openings incl. a
//     "Canada - ON - TOR - Remote" filter at
//     https://www.pega.com/about/careers/job-listings — but it sits behind
//     bot protection this sweep does not attempt to defeat.
//   Workable-hosted boards (float.com among them) — public JSON returns the
//     account shell but zero postings for every account tested; the board is
//     a client-side app with no server-rendered payload to scrape. Not a
//     supported ATS here; check by hand if that vendor matters to you.
//   Workday tenants beyond Clio/Zendesk need a 'tenant:host:site' triple read
//     out of a browser address bar — the site id can't be brute-forced.

/* ------------------------------------------------------------------ */
/* Title / location classification                                    */
/* ------------------------------------------------------------------ */

// Two-part test, because a bare "Senior Manager" matches everything from
// Executive Compensation to Public Affairs. A title qualifies if it is
// either unambiguously an engineering-leadership title (STRONG), or it
// carries a leadership level (LEVEL) *and* an engineering signal (ENG_SIGNAL).
const STRONG =
  /\b(engineering manager|director of engineering|head of engineering|vp,?\s+engineering|vice president,?\s+engineering|manager,?\s+(software|engineering)|(software|development)\s+manager|engineering\s+director)\b/i;

const LEVEL =
  /\b(manager|senior manager|sr\.?\s+manager|director|senior director|sr\.?\s+director|head of|vp|vice president)\b/i;

const ENG_SIGNAL =
  /\b(engineering|software|development|developer|platform|technical|technology|infrastructure|architecture|devops|sre|reliability|backend|back-end|frontend|front-end|fullstack|full-stack|mobile|api|data platform|data engineering|machine learning|ml\b|ai engineering|security engineering|qa\b|test engineering|cloud|systems)\b/i;

// Titles that carry an engineering signal but are not engineering leadership.
// NOTE: prefix stems (no trailing \b) — a trailing \b silently fails on
// inflected words, e.g. \brecruit\b never matches "Recruiting".
const NOT_ENGINEERING =
  /\b(product manager|product marketing|group product|program manager|project manager|account|customer success|customer experience|sales|revenue|renewal|marketing|recruit|talent|people|payroll|compensat|benefits|procurement|financ|fp&a|accounting|legal|counsel|polic|public affairs|communicat|press|brand|support|solutions architect|field engineering|sales engineering|professional services|consultant|partner|channel|alliance|community|content|design|\bux\b|user research|analyst|analytics manager|information technology|it operations|governance|risk|audit|compliance|trust & safety|business development|corporate strategy|operations manager|office manager|facilities|technical program|technical account|implementation|onboarding|enablement|training|education|technical learning|learning &|learning and development|business systems|corporate systems|internal systems|\babm\b|\bhr\b|\btax\b)/i;

function isEngLeadership(title: string): boolean {
  if (NOT_ENGINEERING.test(title)) return false;
  if (STRONG.test(title)) return true;
  return LEVEL.test(title) && ENG_SIGNAL.test(title);
}

/** Also qualify a title if it contains one of the user's own configured
 *  target titles verbatim — widens the net past the built-in classifier for
 *  a title it wouldn't otherwise recognize (e.g. a niche "Head of Platform"
 *  variant), without weakening the classifier itself. */
function matchesOwnTitles(title: string, jobTitles: string[]): boolean {
  const t = title.toLowerCase();
  return jobTitles.some((jt) => jt.trim() && t.includes(jt.trim().toLowerCase()));
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Well-tuned Canada synonyms (metro nicknames, major cities) — kept as the
// default when the saved criteria's country is Canada, since a from-scratch
// generic match would miss "GTA" / "Mississauga" / etc.
const CANADA_OK =
  /(canada|canadian|toronto|ontario|\bgta\b|mississauga|markham|waterloo|kitchener|ottawa|montreal|vancouver|calgary)/i;
const CANADA_MAYBE =
  /(north america|namer|americas|remote\s*[-,]?\s*global|worldwide|anywhere|emea\/namer)/i;
const NOT_CANADA =
  /(united states|\bu\.?s\.?a?\b|remote\s*[-,]?\s*us\b|us remote|india|bangalore|bengaluru|poland|warsaw|germany|berlin|munich|france|paris|spain|madrid|ireland|dublin|united kingdom|london|netherlands|amsterdam|australia|sydney|singapore|japan|tokyo|israel|tel aviv|brazil|mexico|costa rica|bulgaria|cyprus|nicosia|prague|czech|reykjav|austria|vienna|new york|san francisco|seattle|bellevue|austin|chicago|boston|denver|los angeles|atlanta|san jose|washington)/i;

interface LocationMatchers {
  ok: RegExp;
  maybe: RegExp;
  not: RegExp;
}

/** Build the location-qualification regexes from the saved criteria. Canada
 *  gets the hand-tuned synonym list above; any other country falls back to
 *  matching the country/city name itself plus the same broad-remote-scope
 *  "maybe" bucket — looser, but there's no tuned exclude-list to pair with
 *  an arbitrary country, so nothing is hard-excluded in that case. */
function buildLocationMatchers(criteria: SearchCriteria): LocationMatchers {
  const country = (criteria.country || "Canada").trim();
  if (/^can(ada)?$/i.test(country)) {
    return { ok: CANADA_OK, maybe: CANADA_MAYBE, not: NOT_CANADA };
  }
  const terms = [country, criteria.city].filter((s): s is string => !!s && s.trim() !== "");
  const ok = terms.length
    ? new RegExp(`(${terms.map(escapeRegExp).join("|")})`, "i")
    : /(?!)/; // matches nothing if neither country nor city is set
  return { ok, maybe: CANADA_MAYBE, not: /(?!)/ };
}

function locationBucket(locText: string, m: LocationMatchers): "canada" | "maybe" | "no" | "unknown" {
  if (!locText) return "unknown";
  if (m.ok.test(locText)) return "canada"; // "qualifies", regardless of country
  if (m.maybe.test(locText) && !m.not.test(locText)) return "maybe";
  if (m.not.test(locText)) return "no";
  return "unknown";
}

/* ------------------------------------------------------------------ */
/* Compensation text mining                                           */
/* ------------------------------------------------------------------ */

const MONEY = /(?:CA?\$|\$|CAD|USD)\s?(\d{2,3}(?:,\d{3})+|\d{6})/g;

function moneyList(text: string): number[] {
  if (!text) return [];
  const out: number[] = [];
  for (const m of text.matchAll(MONEY)) out.push(parseInt(m[1].replace(/,/g, ""), 10));
  return out;
}

function currencyOf(s: string): string {
  if (/\bCA\$|\bCAD\b|\bcanad/i.test(s)) return "CAD";
  if (/\bUSD\b|\bUS\$|\bu\.?s\.?\b/i.test(s)) return "USD";
  return "";
}

interface CompResult {
  snippets: string[];
  currency: string;
  low: number | null;
  high: number | null;
}

interface AshbyCompTier {
  tierSummary?: string;
}
interface AshbyComp {
  compensationTiers?: AshbyCompTier[];
}

/**
 * Postings routinely publish a US band AND a Canadian band in the same body.
 * Taking min/max across the whole text would invent a range that doesn't
 * exist, so return the verbatim sentences around each figure and derive the
 * range from the single best snippet (Canada-specific if there is one).
 */
function compSnippets(text: string | undefined, ashbyComp: AshbyComp | null | undefined): CompResult {
  const snips: string[] = [];
  if (ashbyComp?.compensationTiers) {
    for (const tier of ashbyComp.compensationTiers) {
      const label = (tier.tierSummary || "").trim();
      if (label && moneyList(label).length) snips.push(label);
    }
  }
  const body = (text || "").replace(/\s+/g, " ");
  for (const m of body.matchAll(MONEY)) {
    const start = Math.max(0, (m.index ?? 0) - 130);
    const end = (m.index ?? 0) + m[0].length + 90;
    const frag = body.slice(start, end).trim();
    if (!snips.some((s) => frag.slice(0, 60).includes(s.slice(0, 60)) || s.slice(0, 60).includes(frag.slice(0, 60)))) {
      snips.push(frag);
    }
    if (snips.length >= 8) break;
  }
  if (!snips.length) return { snippets: [], currency: "", low: null, high: null };
  const canadian = snips.filter((s) => currencyOf(s) === "CAD");
  const best = canadian[0] ?? snips[0];
  const vals = moneyList(best);
  return {
    snippets: snips.slice(0, 3).map((s) => s.slice(0, 190)),
    currency: currencyOf(best),
    low: vals.length ? Math.min(...vals) : null,
    high: vals.length ? Math.max(...vals) : null,
  };
}

/* ------------------------------------------------------------------ */
/* HTML cleanup (Greenhouse/Workday descriptions are HTML, sometimes    */
/* HTML-escaped twice)                                                 */
/* ------------------------------------------------------------------ */

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ",
  copy: "©", reg: "®", mdash: "—", ndash: "–",
  hellip: "…", rsquo: "’", lsquo: "‘", rdquo: "”",
  ldquo: "“", trade: "™", bull: "•", middot: "·",
};

function decodeHtmlEntities(s: string): string {
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, ent: string) => {
    if (ent[0] === "#") {
      const isHex = ent[1] === "x" || ent[1] === "X";
      const code = parseInt(ent.slice(isHex ? 2 : 1), isHex ? 16 : 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : whole;
    }
    return NAMED_ENTITIES[ent] ?? whole;
  });
}

function stripHtml(s: string): string {
  let out = s;
  for (let i = 0; i < 2; i++) {
    out = decodeHtmlEntities(out);
    out = out.replace(/<[^>]+>/g, " ");
  }
  return out.replace(/\s+/g, " ");
}

/* ------------------------------------------------------------------ */
/* HTTP                                                                */
/* ------------------------------------------------------------------ */

interface HttpResult {
  status: number;
  body: string;
}

/* c8 ignore start -- exercising this means a real outbound HTTP call to a
 * live ATS endpoint; the test suite deliberately never does that (see
 * README's Testing section). Everything below down to the matching "c8
 * ignore stop" is a function whose entire body is that live call — the pure
 * logic around it (title/location/salary matching, board-response parsing,
 * candidate ranking) is NOT excluded and stays a real, visible gap. */
/** GET or POST a URL; never throws — network failures come back as status 0. */
async function httpFetch(
  url: string,
  opts: { method?: "GET" | "POST"; body?: string; timeoutMs?: number; retries?: number } = {}
): Promise<HttpResult> {
  const { method = "GET", body, timeoutMs = 25_000, retries = 2 } = opts;
  let lastErr = "";
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const res = await fetch(url, {
        method,
        headers: {
          "User-Agent": UA,
          Accept: "application/json",
          ...(body ? { "Content-Type": "application/json" } : {}),
        },
        body,
        signal: AbortSignal.timeout(timeoutMs),
      });
      const text = await res.text();
      return { status: res.status, body: text };
    } catch (err: any) {
      lastErr = `${err?.name ?? "Error"}: ${err?.message ?? String(err)}`;
    }
  }
  return { status: 0, body: lastErr };
}
/* c8 ignore stop */

const BOARD_URL: Partial<Record<Ats, string>> = {
  greenhouse: "https://boards-api.greenhouse.io/v1/boards/%s/jobs",
  ashby: "https://api.ashbyhq.com/posting-api/job-board/%s?includeCompensation=true",
  lever: "https://api.lever.co/v0/postings/%s?mode=json",
  smartrecruiters: "https://api.smartrecruiters.com/v1/companies/%s/postings?limit=100",
  // workday has no single GET URL — POST-only and paged, see fetchWorkday().
};

const WD_LIMIT = 20; // Workday rejects limit > 20 with HTTP 400
const WD_MAX_PAGES = 15; // 300 postings/board; well past what the title filter needs

function wdParts(token: string): { tenant: string; host: string; site: string } {
  const [tenant, host, site] = token.split(":");
  return { tenant, host, site };
}

function wdCxs(token: string): string {
  const { tenant, host, site } = wdParts(token);
  return `https://${tenant}.${host}.myworkdayjobs.com/wday/cxs/${tenant}/${site}`;
}

/* c8 ignore start -- both functions below make a live outbound HTTP call
 * (see the httpFetch note above); the pure wdParts()/wdCxs() helpers above
 * and parseBoard() below are deliberately NOT excluded. */
/** Page the CXS jobs endpoint into one synthetic board payload. Three
 *  Workday quirks drive this: POST-only, `limit` capped at 20, and the list
 *  view reports dates as relative prose — the real posted date has to come
 *  from the per-posting detail call in workdayDetail(). */
async function fetchWorkday(token: string): Promise<HttpResult> {
  const base = wdCxs(token);
  const items: any[] = [];
  let total: number | null = null;
  let status = 0;
  for (let page = 0; page < WD_MAX_PAGES; page++) {
    const r = await httpFetch(`${base}/jobs`, {
      method: "POST",
      body: JSON.stringify({
        appliedFacets: {},
        limit: WD_LIMIT,
        offset: page * WD_LIMIT,
        searchText: "",
      }),
    });
    status = r.status;
    if (status !== 200) break;
    let j: any;
    try {
      j = JSON.parse(r.body);
    } catch {
      break;
    }
    const batch: any[] = j.jobPostings || [];
    items.push(...batch);
    if (total === null) total = j.total ?? null;
    if (batch.length < WD_LIMIT || (total !== null && items.length >= total)) break;
  }
  if (!items.length) return { status: status || 0, body: "" };
  return { status: 200, body: JSON.stringify({ jobPostings: items }) };
}

async function fetchBoard(ats: Ats, token: string): Promise<HttpResult> {
  if (ats === "workday") return fetchWorkday(token);
  const template = BOARD_URL[ats];
  if (!template) return { status: 0, body: `unsupported ats: ${ats}` };
  return httpFetch(template.replace("%s", token));
} /* c8 ignore stop */

/* ------------------------------------------------------------------ */
/* Per-ATS parsing                                                     */
/* ------------------------------------------------------------------ */

export interface Posting {
  company: string;
  ats: Ats;
  token: string;
  jid: string;
  title: string;
  location: string;
  url: string;
  published: string; // "YYYY-MM-DD" or ""
  compText: string;
  ashbyComp: AshbyComp | null;
  needsDetail: boolean;
}

function parseBoard(company: string, ats: Ats, token: string, body: string): Posting[] {
  let data: any;
  try {
    data = JSON.parse(body);
  } catch {
    return [];
  }
  const out: Posting[] = [];

  if (ats === "greenhouse") {
    for (const j of data.jobs || []) {
      const loc = j.location?.name || "";
      const offices = (j.offices || []).map((o: any) => o.name || "").join(" ");
      out.push({
        company, ats, token, jid: String(j.id),
        title: j.title || "", location: `${loc} ${offices}`.trim(),
        url: j.absolute_url || "",
        published: (j.first_published || j.updated_at || "").slice(0, 10),
        compText: "", ashbyComp: null, needsDetail: true,
      });
    }
  } else if (ats === "ashby") {
    for (const j of data.jobs || []) {
      let loc = j.location || "";
      const sec = (j.secondaryLocations || []).map((s: any) => s.location || "").join(", ");
      if (j.isRemote) loc = "Remote " + loc;
      out.push({
        company, ats, token, jid: String(j.id),
        title: j.title || "", location: `${loc} ${sec}`.trim(),
        url: j.jobUrl || j.applyUrl || "",
        published: (j.publishedAt || "").slice(0, 10),
        compText: (j.descriptionPlain || "").slice(0, 6000),
        ashbyComp: j.compensation || null, needsDetail: false,
      });
    }
  } else if (ats === "lever") {
    for (const j of Array.isArray(data) ? data : []) {
      const cats = j.categories || {};
      let published = "";
      if (j.createdAt) {
        published = new Date(j.createdAt).toISOString().slice(0, 10);
      }
      out.push({
        company, ats, token, jid: String(j.id),
        title: j.text || "", location: cats.location || "",
        url: j.hostedUrl || "",
        published,
        compText: (j.descriptionPlain || "").slice(0, 6000),
        ashbyComp: null, needsDetail: false,
      });
    }
  } else if (ats === "workday") {
    const { tenant, host, site } = wdParts(token);
    const siteBase = `https://${tenant}.${host}.myworkdayjobs.com/${site}`;
    for (const j of data.jobPostings || []) {
      const extPath = j.externalPath || "";
      out.push({
        company, ats, token,
        jid: extPath, // doubles as id and the detail-call path
        title: j.title || "", location: j.locationsText || "",
        url: siteBase + extPath,
        published: "", // list view has relative prose only
        compText: "", ashbyComp: null, needsDetail: true,
      });
    }
  } else if (ats === "smartrecruiters") {
    for (const j of data.content || []) {
      const loc = j.location || {};
      let locText = ["city", "region", "country"].map((k) => String(loc[k] ?? "")).join(" ");
      if (loc.remote) locText = "Remote " + locText;
      out.push({
        company, ats, token, jid: String(j.id),
        title: j.name || "", location: locText.trim(),
        url: `https://jobs.smartrecruiters.com/${token}/${j.id}`,
        published: (j.releasedDate || "").slice(0, 10),
        compText: "", ashbyComp: null, needsDetail: false,
      });
    }
  }
  return out;
}

/* c8 ignore start -- all three functions below make a live outbound HTTP
 * call (see the httpFetch note above). */
/** Fill comp_text/published for a surviving Greenhouse candidate. */
async function greenhouseDetail(p: Posting): Promise<Posting> {
  const url = `https://boards-api.greenhouse.io/v1/boards/${p.token}/jobs/${p.jid}`;
  const { status, body } = await httpFetch(url, { timeoutMs: 20_000, retries: 1 });
  if (status !== 200) return p;
  let d: any;
  try {
    d = JSON.parse(body);
  } catch {
    return p;
  }
  p.compText = stripHtml(d.content || "").slice(0, 8000);
  if (d.first_published) p.published = String(d.first_published).slice(0, 10);
  for (const loc of d.offices || []) {
    p.location = `${p.location} ${loc.name || ""}`.trim();
  }
  return p;
}

/** Fill comp_text/published/location for a surviving Workday candidate — not
 *  optional the way Greenhouse's is: the list view carries no real posted
 *  date and no description at all. */
async function workdayDetail(p: Posting): Promise<Posting> {
  const { status, body } = await httpFetch(wdCxs(p.token) + p.jid, { timeoutMs: 20_000, retries: 1 });
  if (status !== 200) return p;
  let info: any;
  try {
    info = JSON.parse(body).jobPostingInfo || {};
  } catch {
    return p;
  }
  p.compText = stripHtml(info.jobDescription || "").slice(0, 8000);
  if (info.startDate) p.published = String(info.startDate).slice(0, 10);
  const extra = [info.location || "", info.country?.descriptor || ""];
  p.location = [p.location, ...extra].filter(Boolean).join(" ").trim();
  if (info.externalUrl) p.url = info.externalUrl;
  return p;
}

async function fillDetail(p: Posting): Promise<Posting> {
  return p.ats === "workday" ? workdayDetail(p) : greenhouseDetail(p);
} /* c8 ignore stop */

export type LiveStatus = "GONE" | "ok" | "blocked" | "unknown";

/**
 * HTTP-check the public posting URL, deliberately asymmetric:
 *   404/410 -> GONE (the only statuses that actually prove a pull).
 *   403/401/429 -> 'blocked' (bot protection; says nothing about openness).
 *   200 -> 'ok', but weakly — a client-rendered career site returns 200 for
 *     a pulled posting too, so this is "not provably gone", not "confirmed".
 */
/* c8 ignore start -- makes a live outbound HTTP call to the posting URL
 * (see the httpFetch note above); its status->live classification is pure
 * but entangled with the fetch in one small function, so it goes too. */
async function checkLive(p: Posting & { liveStatus?: number; live?: LiveStatus }): Promise<void> {
  const { status } = await httpFetch(p.url, { timeoutMs: 15_000, retries: 0 });
  p.liveStatus = status;
  if (status === 404 || status === 410) p.live = "GONE";
  else if (status === 200) p.live = "ok";
  else if (status === 401 || status === 403 || status === 429) p.live = "blocked";
  else p.live = "unknown";
}
/* c8 ignore stop */

/* ------------------------------------------------------------------ */
/* Concurrency helper                                                  */
/* ------------------------------------------------------------------ */

async function mapLimit<T>(items: T[], limit: number, fn: (item: T) => Promise<void>): Promise<void> {
  let idx = 0;
  async function worker(): Promise<void> {
    for (;;) {
      const i = idx++;
      if (i >= items.length) return;
      await fn(items[i]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
}

/* ------------------------------------------------------------------ */
/* Dedupe against the spreadsheets                                    */
/* ------------------------------------------------------------------ */

interface Existing {
  discoveryPairs: Set<string>;
  discoveryLinks: Set<string>;
  trackerPairs: Set<string>;
  trackerCompanyRows: Map<string, string[]>;
}

function readExisting(): Existing {
  const discoveryPairs = new Set<string>();
  const discoveryLinks = new Set<string>();
  const trackerPairs = new Set<string>();
  const trackerCompanyRows = new Map<string, string[]>();

  if (fs.existsSync(DISCOVERY.filePath)) {
    const h = openWorkbook(DISCOVERY);
    for (const rec of readAllRecords(h)) {
      const company = rec.values["Company"];
      if (!company) continue;
      discoveryPairs.add(keyOf(company, rec.values["Position"] ?? ""));
      for (const l of linksIn(rec.values["Job Link"])) discoveryLinks.add(l);
    }
  }
  if (fs.existsSync(TRACKER.filePath)) {
    const h = openWorkbook(TRACKER);
    for (const rec of readAllRecords(h)) {
      const company = rec.values["Company"];
      if (!company) continue;
      const position = rec.values["Position"] ?? "";
      trackerPairs.add(keyOf(company, position));
      const status = rec.values["Status"] || "no status";
      const nc = norm(company);
      const rows = trackerCompanyRows.get(nc) ?? [];
      rows.push(`${position} (${status})`);
      trackerCompanyRows.set(nc, rows);
      for (const l of linksIn(rec.values["Job Link"])) discoveryLinks.add(l);
    }
  }
  return { discoveryPairs, discoveryLinks, trackerPairs, trackerCompanyRows };
}

/* ------------------------------------------------------------------ */
/* Board-health cache                                                  */
/* ------------------------------------------------------------------ */

interface BoardHealthEntry {
  ok: boolean;
  status: number;
  n: number;
  checked: string;
}
type BoardHealth = Record<string, BoardHealthEntry>;

function readHealth(): BoardHealth {
  try {
    return JSON.parse(fs.readFileSync(SWEEP_HEALTH_FILE, "utf-8"));
  } catch {
    return {};
  }
}

function writeHealth(health: BoardHealth): void {
  try {
    fs.writeFileSync(SWEEP_HEALTH_FILE, JSON.stringify(health, null, 1));
  } catch {
    /* best-effort cache; not fatal */
  }
}

/* ------------------------------------------------------------------ */
/* Candidate shape + ranking                                           */
/* ------------------------------------------------------------------ */

export interface Candidate extends Posting {
  locBucket: "canada" | "maybe" | "no" | "unknown";
  compSnippets: string[];
  currency: string;
  compLow: number | null;
  compHigh: number | null;
  ageDays: number | null;
  inDiscovery: boolean;
  inTracker: boolean;
  companyTrackedAs: string[];
  isNew: boolean;
  clearsMinSalary: boolean;
  liveStatus?: number;
  live?: LiveStatus;
}

export interface SweepResult {
  date: string;
  counts: {
    boardsOk: number;
    boardsTried: number;
    postings: number;
    leadership: number;
    geoPlausible: number;
    new: number;
  };
  deadBoards: string[];
  candidates: Candidate[]; // geo-plausible, both new and already-known
}

export interface SweepOptions {
  /** Skip boards recorded as unreachable in board_health.json. */
  quick?: boolean;
  /** Concurrent HTTP requests. Default 12, matching the old sweep.py default. */
  workers?: number;
  /** Persist sweep_out.json / board_health.json. Default true. */
  persist?: boolean;
}

function todayISO(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function daysBetween(isoDate: string, today: string): number | null {
  const a = Date.parse(isoDate + "T00:00:00Z");
  const b = Date.parse(today + "T00:00:00Z");
  if (Number.isNaN(a) || Number.isNaN(b)) return null;
  return Math.round((b - a) / 86_400_000);
}

/** Run the full board sweep. Returns geo-plausible candidates (new and
 *  already-known); callers typically only surface the new ones. */
export async function runSweep(criteria: SearchCriteria, opts: SweepOptions = {}): Promise<SweepResult> {
  const { quick = false, workers = 12, persist = true } = opts;
  const today = todayISO();
  const locMatchers = buildLocationMatchers(criteria);
  const minSalary = criteria.minSalary ?? 0;
  const jobTitles = criteria.jobTitles ?? [];

  const health = readHealth();
  const boards = quick
    ? BOARDS.filter((b) => health[`${b[1]}:${b[2]}`]?.ok ?? true)
    : BOARDS;

  let postings: Posting[] = [];
  let boardsOk = 0;
  const deadBoards: string[] = [];

  await mapLimit(boards, workers, async ([name, ats, token]) => {
    const { status, body } = await fetchBoard(ats, token);
    const got = status === 200 ? parseBoard(name, ats, token, body) : [];
    health[`${ats}:${token}`] = { ok: status === 200, status, n: got.length, checked: today };
    if (got.length) {
      boardsOk++;
      postings.push(...got);
    } else {
      deadBoards.push(`${ats}/${token}(${status})`);
    }
  });

  // Filter on title, then location.
  const lead = postings.filter((p) => isEngLeadership(p.title) || matchesOwnTitles(p.title, jobTitles));
  const geoAll: (Posting & Partial<Candidate>)[] = lead.map((p) => ({
    ...p,
    locBucket: locationBucket(p.location, locMatchers),
  }));
  let geo = geoAll.filter((p) => p.locBucket !== "no");

  // Detail fetch only for survivors that need it.
  const need = geo.filter((p) => p.needsDetail);
  if (need.length) await mapLimit(need, workers, async (p) => void (await fillDetail(p as Posting)));

  // Re-bucket: Greenhouse/Workday detail can add office names.
  for (const p of geo) p.locBucket = locationBucket(p.location, locMatchers);
  geo = geo.filter((p) => p.locBucket !== "no");

  // Comp + age + dedupe.
  const existing = readExisting();
  const candidates: Candidate[] = geo.map((p) => {
    const comp = compSnippets(p.compText, p.ashbyComp);
    const ageDays = p.published ? daysBetween(p.published, today) : null;
    const k = keyOf(p.company, p.title);
    const inDiscovery = existing.discoveryPairs.has(k) || existing.discoveryLinks.has(normUrl(p.url));
    const inTracker = existing.trackerPairs.has(k);
    return {
      ...(p as Posting),
      locBucket: p.locBucket as Candidate["locBucket"],
      compSnippets: comp.snippets,
      currency: comp.currency,
      compLow: comp.low,
      compHigh: comp.high,
      ageDays,
      inDiscovery,
      inTracker,
      companyTrackedAs: existing.trackerCompanyRows.get(norm(p.company)) ?? [],
      isNew: !inDiscovery && !inTracker,
      clearsMinSalary: !!(comp.high && minSalary > 0 && comp.high >= minSalary),
      compText: "",
      ashbyComp: null,
    };
  });

  // Liveness: only the new candidates, since those are the only URLs that
  // will be written to the sheet or acted on.
  const fresh = candidates.filter((c) => c.isNew);
  if (fresh.length) await mapLimit(fresh, workers, checkLive);
  const survivors = fresh.filter((c) => c.live !== "GONE");
  const droppedGone = fresh.length - survivors.length;

  function rank(c: Candidate): [number, number, number, number] {
    return [
      c.locBucket === "canada" ? 0 : 1,
      c.currency === "CAD" ? 0 : 1,
      c.clearsMinSalary ? 0 : 1,
      c.ageDays ?? 999,
    ];
  }
  const survivorKeys = new Set(survivors.map((c) => keyOf(c.company, c.title) + "|" + c.url));
  const finalCandidates = candidates.filter(
    (c) => !c.isNew || survivorKeys.has(keyOf(c.company, c.title) + "|" + c.url)
  );
  finalCandidates
    .filter((c) => c.isNew)
    .sort((a, b) => {
      const ra = rank(a);
      const rb = rank(b);
      for (let i = 0; i < ra.length; i++) if (ra[i] !== rb[i]) return ra[i] - rb[i];
      return 0;
    });

  const result: SweepResult = {
    date: today,
    counts: {
      boardsOk,
      boardsTried: boards.length,
      postings: postings.length,
      leadership: lead.length,
      geoPlausible: geo.length,
      new: survivors.length,
    },
    deadBoards: deadBoards.sort(),
    candidates: finalCandidates,
  };

  if (persist) {
    writeHealth(health);
    try {
      fs.writeFileSync(
        SWEEP_OUT_FILE,
        JSON.stringify(
          { ...result, droppedGoneLinks: droppedGone },
          null,
          1
        )
      );
    } catch {
      /* best-effort dump; the in-memory result is still returned */
    }
  }

  return result;
}
