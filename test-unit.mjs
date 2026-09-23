/**
 * Unit tests for pure logic that test-client.mjs can't reach through the
 * MCP tool surface: sweepEngine.ts/gmailTools.ts/gmailAuth.ts helper
 * functions whose only callers make a real network/OAuth call (and are
 * therefore excluded from coverage — see the /* c8 ignore *\/ comments in
 * those files). The functions tested here make no such call themselves, so
 * they're imported directly from dist/ and exercised with plain assertions
 * instead of spawning a server process.
 *
 *   npm run build && node test-unit.mjs
 */
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import * as XLSX from "xlsx";
XLSX.set_fs(fs);

// Set before the dist imports below: config.js resolves JOB_TRACKER_FILE/
// JOB_DISCOVERY_FILE (and derives JOB_ROOT, and so SWEEP_HEALTH_FILE) at
// import time, and readExisting()/readHealth()/writeHealth() actually touch
// the filesystem at those paths — point them at a throwaway temp folder so
// nothing real is ever read or written.
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jobtracker-unit-"));
process.env.JOB_TRACKER_FILE = path.join(dir, "Job_Tracking.xlsx");
process.env.JOB_DISCOVERY_FILE = path.join(dir, "Job_Search_Discovery.xlsx");

const sweep = await import("./dist/sweepEngine.js");
const gmailTools = await import("./dist/gmailTools.js");
const gmailAuth = await import("./dist/gmailAuth.js");
const xlsxFormat = await import("./dist/xlsxFormat.js");
const ExcelJS = (await import("exceljs")).default;

let passed = 0, failed = 0;
const fails = [];
function check(name, cond, extra) {
  if (cond) { passed++; console.log(`  PASS  ${name}`); }
  else { failed++; fails.push(name); console.log(`  FAIL  ${name}${extra ? "  -> " + extra : ""}`); }
}

/* ---- fixture helpers (mirrors test-client.mjs's, kept self-contained) ---- */
function makeSheet(headers, rows) {
  const ws = {};
  headers.forEach((h, c) => (ws[XLSX.utils.encode_cell({ r: 0, c })] = { t: "s", v: h }));
  rows.forEach((row, ri) => {
    headers.forEach((h, c) => {
      const v = row[h];
      if (v === undefined || v === "") return;
      ws[XLSX.utils.encode_cell({ r: ri + 1, c })] = { t: "s", v: String(v) };
    });
  });
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: headers.length - 1 } });
  return ws;
}
function writeBook(file, ws, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, file);
}

/* ==================================================================== */
/* sweepEngine.ts                                                        */
/* ==================================================================== */

console.log("\n# sweepEngine: title classification (isEngLeadership / matchesOwnTitles)");
check("recognizes 'Engineering Manager'", sweep.isEngLeadership("Engineering Manager"));
check("recognizes 'Director of Engineering'", sweep.isEngLeadership("Director of Engineering"));
check("recognizes 'VP, Engineering'", sweep.isEngLeadership("VP, Engineering"));
check("recognizes level+signal combo 'Senior Director, Platform'", sweep.isEngLeadership("Senior Director, Platform"));
check("rejects 'Product Manager'", !sweep.isEngLeadership("Product Manager"));
check("rejects 'Manager, Talent Acquisition'", !sweep.isEngLeadership("Manager, Talent Acquisition"));
check("rejects a bare IC title 'Software Engineer'", !sweep.isEngLeadership("Software Engineer"));
check("rejects manager with no eng signal 'Account Manager'", !sweep.isEngLeadership("Account Manager"));
check(
  "matchesOwnTitles finds a configured title as a substring",
  sweep.matchesOwnTitles("Head of Platform Engineering", ["Head of Platform"])
);
check("matchesOwnTitles is case-insensitive", sweep.matchesOwnTitles("HEAD OF PLATFORM", ["head of platform"]));
check("matchesOwnTitles false when nothing matches", !sweep.matchesOwnTitles("Recruiter", ["Engineering Manager"]));

console.log("\n# sweepEngine: location bucketing");
const caCriteria = { jobTitles: [], workStyle: [], city: "", country: "Canada", minSalary: 0, currency: "", notes: "" };
const caMatchers = sweep.buildLocationMatchers(caCriteria);
check("Toronto buckets as canada", sweep.locationBucket("Toronto, ON", caMatchers) === "canada");
check("Remote (Canada) buckets as canada", sweep.locationBucket("Remote (Canada)", caMatchers) === "canada");
check("a bare 'Remote' with no scope is unknown, not an automatic match", sweep.locationBucket("Remote", caMatchers) === "unknown");
check("'Remote - Global' buckets as maybe", sweep.locationBucket("Remote - Global", caMatchers) === "maybe");
check("United States buckets as no", sweep.locationBucket("United States", caMatchers) === "no");
check("blank location is unknown", sweep.locationBucket("", caMatchers) === "unknown");
check("an unrecognized location is unknown", sweep.locationBucket("Antarctica Base", caMatchers) === "unknown");

const usCriteria = { jobTitles: [], workStyle: [], city: "Austin", country: "United States", minSalary: 0, currency: "", notes: "" };
const usMatchers = sweep.buildLocationMatchers(usCriteria);
check("a non-Canada criteria matches its own country/city", sweep.locationBucket("Austin, TX", usMatchers) === "canada");
check("a non-Canada criteria has no hard-exclude list", sweep.locationBucket("Toronto, ON", usMatchers) !== "no");

console.log("\n# sweepEngine: salary text mining");
check("moneyList parses a single figure", JSON.stringify(sweep.moneyList("Salary: $180,000")) === JSON.stringify([180000]));
check(
  "moneyList parses a range",
  JSON.stringify(sweep.moneyList("CA$150,000 - CA$180,000")) === JSON.stringify([150000, 180000])
);
check("moneyList ignores small numbers (not a salary figure)", sweep.moneyList("Team of 12 engineers").length === 0);
check("currencyOf detects CAD", sweep.currencyOf("CA$150,000") === "CAD");
check("currencyOf detects USD", sweep.currencyOf("$150,000 USD") === "USD");
check("currencyOf defaults to empty when ambiguous", sweep.currencyOf("$150,000") === "");

const comp = sweep.compSnippets("We offer a competitive salary of $150,000 - $180,000 USD plus benefits.", null);
check("compSnippets extracts low/high", comp.low === 150000 && comp.high === 180000);
check("compSnippets detects currency from the snippet", comp.currency === "USD");
const compEmpty = sweep.compSnippets("No compensation info here.", null);
check(
  "compSnippets returns nulls when no figure is found",
  compEmpty.low === null && compEmpty.high === null && compEmpty.snippets.length === 0
);
const ashbyComp = { compensationTiers: [{ tierSummary: "CA$140,000 - CA$160,000 CAD" }] };
const compAshby = sweep.compSnippets(undefined, ashbyComp);
check(
  "compSnippets reads Ashby's structured compensation tiers",
  compAshby.currency === "CAD" && compAshby.low === 140000
);

console.log("\n# sweepEngine: HTML cleanup");
check("decodeHtmlEntities handles named entities", sweep.decodeHtmlEntities("Tea &amp; Biscuits") === "Tea & Biscuits");
check("decodeHtmlEntities handles numeric entities", sweep.decodeHtmlEntities("&#8217;") === "’");
// stripHtml deliberately doesn't trim its ends (it's meant for inline use in
// larger text), so tag boundaries can leave a stray leading/trailing space.
check("stripHtml removes tags and decodes entities", sweep.stripHtml("<p>Hello&nbsp;<b>World</b></p>").trim() === "Hello World");
check("stripHtml collapses whitespace", sweep.stripHtml("<div>a</div><div>b</div>").trim() === "a b");

console.log("\n# sweepEngine: Workday token parsing");
// A real Workday token is "tenant:host-code:site", e.g. "acme:wd5:External" —
// the template below already appends ".myworkdayjobs.com", so the host
// segment is a short code like "wd5", not the literal string "myworkdayjobs".
const wd = sweep.wdParts("acme:wd5:External");
check("wdParts splits tenant:host:site", wd.tenant === "acme" && wd.host === "wd5" && wd.site === "External");
check(
  "wdCxs builds the CXS base URL",
  sweep.wdCxs("acme:wd5:External") === "https://acme.wd5.myworkdayjobs.com/wday/cxs/acme/External"
);

console.log("\n# sweepEngine: per-ATS board response parsing (parseBoard)");
const ghBody = JSON.stringify({
  jobs: [{
    id: 1, title: "Engineering Manager", location: { name: "Toronto" },
    offices: [{ name: "TOR" }], absolute_url: "https://job.example/1",
    first_published: "2026-01-01T00:00:00Z",
  }],
});
const gh = sweep.parseBoard("Acme", "greenhouse", "acme", ghBody);
check("parseBoard(greenhouse) extracts one posting", gh.length === 1 && gh[0].title === "Engineering Manager");
check("parseBoard(greenhouse) joins location + offices", gh[0].location === "Toronto TOR");
check("parseBoard(greenhouse) flags needsDetail", gh[0].needsDetail === true);

const ashbyBody = JSON.stringify({
  jobs: [{
    id: "j1", title: "Director of Engineering", location: "Remote", isRemote: true,
    jobUrl: "https://job.example/2", publishedAt: "2026-02-01T00:00:00Z",
    descriptionPlain: "Great role", compensation: { compensationTiers: [] },
  }],
});
const ashby = sweep.parseBoard("Acme", "ashby", "acme", ashbyBody);
check("parseBoard(ashby) prefixes Remote", ashby[0].location.startsWith("Remote"));
check("parseBoard(ashby) does not need a detail call", ashby[0].needsDetail === false);

const leverBody = JSON.stringify([{
  id: "l1", text: "Head of Platform", categories: { location: "Toronto" },
  hostedUrl: "https://job.example/3", createdAt: 1735689600000,
}]);
const lever = sweep.parseBoard("Acme", "lever", "acme", leverBody);
check("parseBoard(lever) reads the array-shaped payload", lever.length === 1 && lever[0].title === "Head of Platform");
check("parseBoard(lever) converts createdAt to an ISO date", /^\d{4}-\d{2}-\d{2}$/.test(lever[0].published));

const wdBody = JSON.stringify({
  jobPostings: [{ title: "Engineering Director", locationsText: "Toronto, ON", externalPath: "/job/123" }],
});
const wdPostings = sweep.parseBoard("Acme", "workday", "acme:wd5:External", wdBody);
check(
  "parseBoard(workday) builds the full URL from siteBase + externalPath",
  wdPostings[0].url === "https://acme.wd5.myworkdayjobs.com/External/job/123"
);
check("parseBoard(workday) flags needsDetail (no posted date in the list view)", wdPostings[0].needsDetail === true);

const srBody = JSON.stringify({
  content: [{
    id: 99, name: "Engineering Manager",
    location: { city: "Toronto", region: "ON", country: "Canada", remote: true },
    releasedDate: "2026-03-01T00:00:00Z",
  }],
});
const sr = sweep.parseBoard("Acme", "smartrecruiters", "acme", srBody);
check(
  "parseBoard(smartrecruiters) prefixes Remote and joins location parts",
  sr[0].location.startsWith("Remote") && sr[0].location.includes("Toronto")
);

check(
  "parseBoard returns [] for malformed JSON rather than throwing",
  sweep.parseBoard("Acme", "greenhouse", "acme", "not json").length === 0
);

console.log("\n# sweepEngine: dates");
check("todayISO returns YYYY-MM-DD", /^\d{4}-\d{2}-\d{2}$/.test(sweep.todayISO()));
check("daysBetween computes a positive day count", sweep.daysBetween("2026-01-01", "2026-01-11") === 10);
check("daysBetween returns null for an unparseable date", sweep.daysBetween("not-a-date", "2026-01-11") === null);

console.log("\n# sweepEngine: mapLimit concurrency helper");
{
  const seen = [];
  let inFlight = 0, maxInFlight = 0;
  await sweep.mapLimit([1, 2, 3, 4, 5], 2, async (item) => {
    inFlight++;
    maxInFlight = Math.max(maxInFlight, inFlight);
    await new Promise((r) => setTimeout(r, 5));
    seen.push(item);
    inFlight--;
  });
  check("mapLimit processes every item exactly once", seen.length === 5 && new Set(seen).size === 5);
  check("mapLimit respects the concurrency limit", maxInFlight <= 2, `maxInFlight=${maxInFlight}`);
}

console.log("\n# sweepEngine: board-health cache round-trip (local file, no network)");
check("readHealth returns {} before any cache file exists", JSON.stringify(sweep.readHealth()) === "{}");
{
  const health = { "greenhouse:acme": { ok: true, status: 200, n: 3, checked: "2026-01-01" } };
  sweep.writeHealth(health);
  check("writeHealth + readHealth round-trips", JSON.stringify(sweep.readHealth()) === JSON.stringify(health));
}

console.log("\n# sweepEngine: readExisting (dedupe source) against real fixture files");
{
  const emptyExisting = sweep.readExisting();
  check(
    "readExisting returns empty sets before the tracker/discovery files exist",
    emptyExisting.trackerPairs.size === 0 && emptyExisting.discoveryPairs.size === 0
  );

  const TRACKER_COLS = [
    "Company", "Position", "Job Link", "Location", "Resume Version",
    "Contact/Referral", "Date Applied", "Status", "Next Follow-Up", "Notes",
  ];
  const DISCOVERY_COLS = [
    "Date Found", "Company", "Position", "Location (Remote/Hybrid)", "Salary",
    "Salary Confidence (Confirmed/Estimated)", "Job Link",
    "Source (Company Careers Page/LinkedIn/Aggregator)",
    "Posted Date / Days Since Posted", "Known Gap Flag", "Match Assessment", "Status",
  ];
  writeBook(
    process.env.JOB_TRACKER_FILE,
    makeSheet(TRACKER_COLS, [
      { Company: "Acme Corp", Position: "Engineering Manager", Status: "Applied", "Job Link": "https://acme.example/jobs/1" },
    ]),
    "Tracking"
  );
  writeBook(
    process.env.JOB_DISCOVERY_FILE,
    makeSheet(DISCOVERY_COLS, [
      { Company: "Globex", Position: "Director of Engineering", "Job Link": "https://globex.example/jobs/2" },
    ]),
    "Discovery"
  );

  const existing = sweep.readExisting();
  check("readExisting picks up a tracker row as a tracked pair", existing.trackerPairs.size === 1);
  check("readExisting picks up a discovery row as a discovery pair", existing.discoveryPairs.size === 1);
  check(
    "readExisting collects Job Link URLs from both sheets",
    existing.discoveryLinks.has("https://acme.example/jobs/1") && existing.discoveryLinks.has("https://globex.example/jobs/2")
  );
  check(
    "readExisting tracks which roles a company already has open (normalized key)",
    existing.trackerCompanyRows.has("acme corp") &&
      existing.trackerCompanyRows.get("acme corp")[0].includes("Engineering Manager")
  );
}

/* ==================================================================== */
/* gmailTools.ts                                                         */
/* ==================================================================== */

console.log("\n# gmailTools: small pure helpers");
check(
  "messageLink builds a Gmail web URL",
  gmailTools.messageLink("abc123") === "https://mail.google.com/mail/u/0/#all/abc123"
);
check(
  "header() finds a header case-insensitively",
  gmailTools.header({ headers: [{ name: "Subject", value: "Hello" }] }, "subject") === "Hello"
);
check(
  "header() returns undefined for a missing header",
  gmailTools.header({ headers: [{ name: "Subject", value: "Hello" }] }, "From") === undefined
);
check(
  "extractEmailAddress pulls the bracketed address out of a display name",
  gmailTools.extractEmailAddress("Jane Recruiter <jane@example.com>") === "jane@example.com"
);
check(
  "extractEmailAddress passes through a bare address unchanged",
  gmailTools.extractEmailAddress("jane@example.com") === "jane@example.com"
);
check("guessStatusFromText detects an offer", gmailTools.guessStatusFromText("We are pleased to offer you the role") === "Offer");
check(
  "guessStatusFromText detects a decline",
  gmailTools.guessStatusFromText("Unfortunately we have decided to move forward with other candidates") === "Declined"
);
check(
  "guessStatusFromText detects an interview request",
  gmailTools.guessStatusFromText("Can we schedule an interview next week?") === "Interviewing"
);
check("guessStatusFromText returns undefined for neutral text", gmailTools.guessStatusFromText("Thanks for your application") === undefined);

console.log("\n# gmailTools: base64url round-trip + MIME body extraction");
const rawText = "Hello, Wörld! 🎉";
const encoded = gmailTools.toBase64Url(rawText);
check("toBase64Url output has no +, /, or = padding", !/[+/=]/.test(encoded));
check("decodeBase64Url reverses toBase64Url", gmailTools.decodeBase64Url(encoded) === rawText);

check(
  "extractBody reads a single-part text/plain message",
  gmailTools.extractBody({ mimeType: "text/plain", body: { data: gmailTools.toBase64Url("plain body") } }).text === "plain body"
);
const multipart = {
  mimeType: "multipart/alternative",
  parts: [
    { mimeType: "text/plain", body: { data: gmailTools.toBase64Url("plain part") } },
    { mimeType: "text/html", body: { data: gmailTools.toBase64Url("<p>html part</p>") } },
  ],
};
const extracted = gmailTools.extractBody(multipart);
check("extractBody finds the text/plain part in a multipart tree", extracted.text === "plain part");
check("extractBody finds the text/html part in a multipart tree", extracted.html === "<p>html part</p>");

console.log("\n# gmailTools: header encoding + raw email construction");
check("encodeHeaderValue passes ASCII through unchanged", gmailTools.encodeHeaderValue("Plain Subject") === "Plain Subject");
const encodedSubject = gmailTools.encodeHeaderValue("Job Search — New Leads");
check("encodeHeaderValue RFC2047-encodes non-ASCII text", /^=\?UTF-8\?B\?/.test(encodedSubject));

const rawEmail = gmailTools.buildRawEmail({
  to: "someone@example.com",
  subject: "Following up",
  body: "Just checking in.",
});
const decodedEmail = gmailTools.decodeBase64Url(rawEmail);
check("buildRawEmail includes the To header", decodedEmail.includes("To: someone@example.com"));
check("buildRawEmail includes the Subject header", decodedEmail.includes("Subject: Following up"));
check("buildRawEmail includes the body after a blank line", decodedEmail.includes("\r\n\r\nJust checking in."));
check("buildRawEmail omits Cc when not given", !decodedEmail.includes("Cc:"));
const rawEmailWithCc = gmailTools.decodeBase64Url(
  gmailTools.buildRawEmail({ to: "a@example.com", cc: "b@example.com", subject: "s", body: "b" })
);
check("buildRawEmail includes Cc when given", rawEmailWithCc.includes("Cc: b@example.com"));

/* ==================================================================== */
/* gmailAuth.ts                                                          */
/* ==================================================================== */

console.log("\n# gmailAuth: pure helpers");
check(
  "buildUrl appends simple params",
  gmailAuth.buildUrl("/messages", { q: "from:x" }) === "https://gmail.googleapis.com/gmail/v1/users/me/messages?q=from%3Ax"
);
check(
  "buildUrl repeats array-valued params (Gmail's metadataHeaders convention)",
  gmailAuth.buildUrl("/messages/1", { metadataHeaders: ["From", "Subject"] }) ===
    "https://gmail.googleapis.com/gmail/v1/users/me/messages/1?metadataHeaders=From&metadataHeaders=Subject"
);
check(
  "buildUrl has no query string when params are empty",
  gmailAuth.buildUrl("/profile", undefined) === "https://gmail.googleapis.com/gmail/v1/users/me/profile"
);

function catchMessage(fn) {
  try { fn(); return null; } catch (e) { return e.message; }
}
check(
  "wrapGmailError flags a 401 as a token problem",
  /re-run/i.test(catchMessage(() => gmailAuth.wrapGmailError({ response: { status: 401 } }, "read /messages")) || "")
);
check(
  "wrapGmailError surfaces the Gmail API's own error message",
  (catchMessage(() => gmailAuth.wrapGmailError(
    { response: { status: 400, data: { error: { message: "Invalid query" } } } },
    "read /messages"
  )) || "").includes("Invalid query")
);

/* ==================================================================== */
/* xlsxFormat.ts                                                        */
/* ==================================================================== */

console.log("\n# xlsxFormat: the SheetJS-then-exceljs handoff must yield a sheet Excel loads without repair");
{
  // Regression coverage for the real "We found a problem with some content"
  // corruption diagnosed from Excel's own recovery log ("sheet1.xml part with
  // XML error. Load error. Line 2, column 0"): SheetJS writes no
  // <sheetFormatPr>, exceljs then reads worksheet.properties back as {} and on
  // write emits <sheetFormatPr customHeight="1"/> WITHOUT the defaultRowHeight
  // attribute the OOXML schema requires. These assertions read the raw part XML
  // (not exceljs's re-parsed model, which would paper over exactly this) so
  // they fail on the bytes Excel actually rejects.
  const JSZip = (await import("jszip")).default;
  const fixturePath = path.join(dir, "Format_fixture.xlsx");
  const HEADERS = ["Date Found", "Company", "Position", "Job Link"];

  // Written the way workbook.ts writes: SheetJS, no `sheetFormat` option.
  function writeSheetJsFixture(rowCount) {
    const aoa = [HEADERS];
    for (let r = 2; r <= rowCount; r++) {
      aoa.push([`2026-01-${String(r).padStart(2, "0")}`, `Co${r}`, `Role${r}`, `https://example.com/jobs/${r}`]);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Discovery");
    XLSX.writeFile(wb, fixturePath);
  }
  async function rawPart(name) {
    const zip = await JSZip.loadAsync(fs.readFileSync(fixturePath));
    return zip.file(name).async("string");
  }
  async function readBack() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(fixturePath);
    return { ws: wb.worksheets[0], definedNames: wb.definedNames.model };
  }
  function sheetFormatPrAttrs(xml) {
    const m = xml.match(/<sheetFormatPr\b([^>]*?)\/?>/);
    return m ? m[1] : null;
  }

  writeSheetJsFixture(10);
  check(
    "precondition: SheetJS itself writes no <sheetFormatPr> (the input exceljs mishandles)",
    !/<sheetFormatPr/.test(await rawPart("xl/worksheets/sheet1.xml"))
  );

  const res = await xlsxFormat.formatWorkbookFile(fixturePath);
  check("formatWorkbookFile succeeds on a SheetJS-written file", res.ok === true, res.detail);
  {
    const xml = await rawPart("xl/worksheets/sheet1.xml");
    const attrs = sheetFormatPrAttrs(xml);
    const height = attrs?.match(/\bdefaultRowHeight="([^"]+)"/);
    check(
      "<sheetFormatPr> carries the schema-required defaultRowHeight (the Excel 'Load error. Line 2' regression)",
      !!height && Number(height[1]) > 0,
      attrs ?? "(no <sheetFormatPr> written)"
    );
    check("<sheetFormatPr> is not flagged customHeight when it just uses Excel's default", !/customHeight="1"/.test(attrs ?? ""), attrs);
    check("<autoFilter> covers the whole table", /<autoFilter ref="A1:D10"\/>/.test(xml));
    check("<sheetViews> freezes the header row", /<pane\b[^>]*\bySplit="1"[^>]*\bstate="frozen"/.test(xml));
    check("<cols> column widths are written", /<cols>.*<col\b[^>]*\bwidth="[^"]+"/.test(xml));
    check(
      "no _xlnm._FilterDatabase is written (exceljs can't scope it; Excel recreates it on save)",
      !/_xlnm\._FilterDatabase/.test(await rawPart("xl/workbook.xml"))
    );

    const { ws } = await readBack();
    const headerCell = ws.getRow(1).getCell(1);
    check("header row is bold", headerCell.font?.bold === true);
    check("header row is filled", headerCell.fill?.fgColor?.argb === "FF305496");
    check("Job Link cells become hyperlinks", typeof ws.getRow(2).getCell(4).value?.hyperlink === "string");
    check("every column gets at least the minimum width", [1, 2, 3, 4].every((c) => (ws.getColumn(c).width ?? 0) >= 12));
  }

  // Real files go through this pass on every sync — a second cycle must stay
  // valid too (exceljs now reads back the <sheetFormatPr> it wrote).
  {
    const res2 = await xlsxFormat.formatWorkbookFile(fixturePath);
    const xml = await rawPart("xl/worksheets/sheet1.xml");
    check("second formatting pass still succeeds", res2.ok === true, res2.detail);
    check("second pass keeps a positive defaultRowHeight", /\bdefaultRowHeight="([1-9][^"]*)"/.test(sheetFormatPrAttrs(xml) ?? ""));
    check("second pass keeps the autoFilter range in sync", /<autoFilter ref="A1:D10"\/>/.test(xml));
  }

  // A file that already declares a valid, non-default row height (e.g. one
  // Excel or openpyxl saved) keeps it, and any _xlnm._FilterDatabase a prior
  // writer left behind is stripped rather than carried forward mis-scoped.
  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Discovery");
    ws.properties.defaultRowHeight = 20;
    ws.addRow(HEADERS);
    ws.addRow(["2026-02-01", "Co", "Role", "https://example.com/x"]);
    wb.definedNames.model = [{ name: "_xlnm._FilterDatabase", ranges: ["Discovery!$A$1:$D$2"] }];
    await wb.xlsx.writeFile(fixturePath);

    const res3 = await xlsxFormat.formatWorkbookFile(fixturePath);
    check("formatting a file with an existing row height succeeds", res3.ok === true, res3.detail);
    const attrs = sheetFormatPrAttrs(await rawPart("xl/worksheets/sheet1.xml"));
    check("an existing valid defaultRowHeight is preserved, not clobbered", /\bdefaultRowHeight="20"/.test(attrs ?? ""), attrs);
    check(
      "a pre-existing _xlnm._FilterDatabase is stripped",
      !/_xlnm\._FilterDatabase/.test(await rawPart("xl/workbook.xml"))
    );
  }
}

fs.rmSync(dir, { recursive: true, force: true });

console.log(`\n${failed === 0 ? "ALL PASSED" : "FAILURES"}: ${passed} passed, ${failed} failed`);
if (failed) {
  console.log("Failed:", fails.join(", "));
  process.exit(1);
}
