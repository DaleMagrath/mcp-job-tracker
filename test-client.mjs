/**
 * Regression test suite for the job-tracker MCP server.
 *
 * Spins up dist/index.js over stdio against throwaway fixtures in a temp folder
 * (your real files are never touched) and exercises every tool with assertions.
 * Prints PASS/FAIL per check and exits non-zero if anything fails.
 *
 *   npm run build && node test-client.mjs
 *
 * Notes:
 *  - Does NOT send a physical print job — print_document is covered only via its
 *    error paths (missing file, unsupported type, unknown printer).
 *  - discovery writes trigger the openpyxl reformat pass; the suite does not
 *    assert on formatting (Python may be absent), only on the data operations.
 */
import { spawn } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import * as XLSX from "xlsx";
XLSX.set_fs(fs);

/* ---- date helpers (Excel serials, TZ-free) ---- */
const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;
const toSerial = (iso) => {
  const [y, m, d] = iso.split("-").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY);
};
const todaySerial = () => {
  const n = new Date();
  return Math.round((Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()) - EPOCH) / DAY);
};
const isoDaysFromToday = (delta) => {
  const ms = EPOCH + (todaySerial() + delta) * DAY;
  const d = new Date(ms);
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(
    d.getUTCDate()
  ).padStart(2, "0")}`;
};

/* ---- fixture builders ---- */
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

function makeSheet(headers, rows, dateCols) {
  const ws = {};
  headers.forEach((h, c) => (ws[XLSX.utils.encode_cell({ r: 0, c })] = { t: "s", v: h }));
  rows.forEach((row, ri) => {
    headers.forEach((h, c) => {
      const v = row[h];
      if (v === undefined || v === "") return;
      const addr = XLSX.utils.encode_cell({ r: ri + 1, c });
      if (dateCols.includes(h)) ws[addr] = { t: "n", v: toSerial(v), z: "d-mmm-yy" };
      else ws[addr] = { t: "s", v: String(v) };
    });
  });
  ws["!ref"] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: rows.length, c: headers.length - 1 },
  });
  return ws;
}
function writeBook(file, ws, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, file);
}

/* ---- JSON-RPC harness ---- */
function startServer(env) {
  const proc = spawn("node", ["dist/index.js"], { env: { ...process.env, ...env } });
  let buf = "";
  const waiters = new Map();
  proc.stdout.on("data", (d) => {
    buf += d;
    let i;
    while ((i = buf.indexOf("\n")) >= 0) {
      const line = buf.slice(0, i);
      buf = buf.slice(i + 1);
      if (!line.trim()) continue;
      let m;
      try { m = JSON.parse(line); } catch { continue; }
      if (m.id && waiters.has(m.id)) { waiters.get(m.id)(m); waiters.delete(m.id); }
    }
  });
  proc.stderr.on("data", () => {});
  let id = 1;
  const rpc = (method, params) =>
    new Promise((res) => {
      const i = ++id;
      waiters.set(i, res);
      proc.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: i, method, params }) + "\n");
    });
  return { proc, rpc };
}

/* ---- assertions ---- */
let passed = 0, failed = 0;
const fails = [];
function check(name, cond, extra) {
  if (cond) { passed++; console.log(`  PASS  ${name}`); }
  else { failed++; fails.push(name); console.log(`  FAIL  ${name}${extra ? "  -> " + extra : ""}`); }
}

async function main() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "jobtracker-test-"));
  const trackerFile = path.join(dir, "Job_Tracking.xlsx");
  const discoveryFile = path.join(dir, "Job_Search_Discovery.xlsx");
  const prepFile = path.join(dir, "Interview_Prep_QA.md");
  const resumesDir = path.join(dir, "Resumes");
  fs.mkdirSync(resumesDir, { recursive: true });

  // Tracker: Amazon (stale + overdue follow-up), Google (fresh, future follow-up).
  writeBook(
    trackerFile,
    makeSheet(
      TRACKER_COLS,
      [
        {
          Company: "Amazon", Position: "Senior Manager",
          "Date Applied": isoDaysFromToday(-20), Status: "Awaiting Response",
          "Next Follow-Up": isoDaysFromToday(-2), Notes: "ping recruiter",
        },
        {
          Company: "Google", Position: "EM",
          "Date Applied": isoDaysFromToday(-5), Status: "Interviewing",
          "Next Follow-Up": isoDaysFromToday(3),
        },
      ],
      ["Date Applied", "Next Follow-Up"]
    ),
    "Tracking"
  );

  // Discovery: one lead to promote.
  writeBook(
    discoveryFile,
    makeSheet(
      DISCOVERY_COLS,
      [
        {
          "Date Found": isoDaysFromToday(-1), Company: "Globex", Position: "EM Payments",
          "Location (Remote/Hybrid)": "Remote (Canada)", Salary: "$220k",
          "Job Link": "https://globex.example/jobs/1", Status: "Open",
          "Match Assessment": "Strong",
        },
      ],
      []
    ),
    "Discovery"
  );

  // Prep: two "Acme" sections to exercise ambiguity handling.
  fs.writeFileSync(
    prepFile,
    "# Interview Prep — Q&A Reference\n\n## Acme — Role A\n\n**Q: Why A?**\n\nA: Because A.\n\n---\n\n## Acme — Role B\n\n**Q: Why B?**\n\nA: Because B.\n"
  );

  // Resume master fixture (stable facts) for generate_resume.
  const masterFile = path.join(dir, "resume_master.json");
  fs.writeFileSync(
    masterFile,
    JSON.stringify({
      name: "Test Candidate",
      contact: "test@example.com • Toronto, ON • linkedin.com/in/test",
      experience: [
        { org: "Acme Corp", title: "Senior Manager", dates: "2020 - Present", description: "Led a team." },
      ],
      education: ["Test University — B.Sc."],
      certifications: ["Test Cert"],
      skills: ["Leadership", "APIs"],
    })
  );

  const { proc, rpc } = startServer({
    JOB_TRACKER_FILE: trackerFile,
    JOB_DISCOVERY_FILE: discoveryFile,
    JOB_INTERVIEW_PREP_FILE: prepFile,
    JOB_RESUMES_DIR: resumesDir,
    JOB_RESUME_MASTER_FILE: masterFile,
    // Point Gmail creds at non-existent files so the suite never touches a real
    // inbox — the Gmail tools should return a clean "authorize first" error.
    JOB_GMAIL_CREDENTIALS_FILE: path.join(dir, "no-credentials.json"),
    JOB_GMAIL_TOKEN_FILE: path.join(dir, "no-token.json"),
  });
  const call = (n, a) =>
    rpc("tools/call", { name: n, arguments: a }).then((m) => {
      const t = m.result?.content?.[0]?.text || "";
      let j; try { j = JSON.parse(t); } catch { j = { message: t }; }
      return { err: !!m.result?.isError, j, text: t };
    });

  await rpc("initialize", {
    protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "test", version: "0" },
  });
  proc.stdin.write(JSON.stringify({ jsonrpc: "2.0", method: "notifications/initialized" }) + "\n");

  const tools = (await rpc("tools/list")).result.tools.map((t) => t.name);

  console.log("\n# Tool registration");
  const expected = [
    "list_jobs", "get_job", "add_job", "update_job_status", "update_job", "delete_job",
    "get_stale_jobs", "get_due_followups", "search_jobs", "draft_followup_message",
    "add_column", "modify_column", "delete_column",
    "discovery_list", "discovery_add", "discovery_update", "discovery_delete", "promote_to_tracker",
    "read_interview_prep", "append_interview_prep",
    "save_document", "print_document", "list_documents", "delete_document",
    "generate_resume", "read_document", "read_text_file",
    "get_resume_master_field", "update_resume_master_field", "list_resume_master_structure",
    "search_gmail_for_job", "read_gmail_message", "scan_job_updates",
    "draft_gmail_reply", "send_gmail_email",
    "get_search_criteria", "update_search_criteria",
    "run_job_sweep", "get_broad_search_queries", "discovery_sync",
    "init_job_tracker_files", "create_resume_master", "check_setup",
  ];
  for (const t of expected) check(`registered: ${t}`, tools.includes(t));

  console.log("\n# Tracker");
  let r = await call("list_jobs", {});
  check("list_jobs count = 2", r.j.count === 2, `got ${r.j.count}`);
  r = await call("get_stale_jobs", { days: 14 });
  check("get_stale_jobs finds Amazon", r.j.count === 1 && r.j.staleJobs[0].Company === "Amazon");
  r = await call("get_due_followups", {});
  check("get_due_followups finds Amazon (overdue)", r.j.count === 1 && r.j.dueFollowUps[0].Company === "Amazon");
  check("get_due_followups daysUntil negative", (r.j.dueFollowUps[0]?.daysUntil ?? 0) < 0);
  r = await call("get_due_followups", { days: 7 });
  check("get_due_followups look-ahead includes Google", r.j.count === 2);
  r = await call("add_job", { company: "Netflix", position: "Director", status: "Applied" });
  check("add_job ok", !r.err);
  r = await call("list_jobs", {});
  check("list_jobs count = 3 after add", r.j.count === 3);
  r = await call("update_job_status", { company: "Netflix", position: "Director", status: "Interviewing" });
  check("update_job_status ok", !r.err);
  r = await call("search_jobs", { query: "recruiter" });
  check("search_jobs finds Amazon by notes", r.j.count === 1 && r.j.jobs[0].Company === "Amazon");
  r = await call("draft_followup_message", { company: "Amazon" });
  check("draft_followup_message returns body", typeof r.j.body === "string" && r.j.body.length > 20);

  console.log("\n# Dynamic columns + get_job resume link");
  r = await call("add_column", { name: "Salary Range", default_value: "TBD" });
  check("add_column ok", !r.err);
  r = await call("update_job", { company: "Amazon", position: "Senior Manager", extra_fields: { "Salary Range": "180-210k" } });
  check("update_job extra_fields ok", !r.err);
  const resumeName = "Dale-Resume-Amazon.pdf";
  await call("save_document", { filename: resumeName, content_base64: Buffer.from("%PDF-1.4 test").toString("base64") });
  r = await call("get_job", { company: "Amazon" });
  check("get_job shows custom column", r.j["Salary Range"] === "180-210k");
  check("get_job links resume file", Array.isArray(r.j.resumeFiles) && r.j.resumeFiles.includes(resumeName));
  r = await call("modify_column", { name: "Salary Range", new_name: "Comp" });
  check("modify_column rename ok", !r.err);
  r = await call("delete_column", { name: "Comp" });
  check("delete_column ok", !r.err);
  r = await call("delete_column", { name: "Company" });
  check("delete_column refuses base column", r.err);

  console.log("\n# Discovery + promote");
  r = await call("discovery_list", {});
  check("discovery_list count = 1", r.j.count === 1);
  r = await call("discovery_add", { company: "Initech", position: "EM", fields: { Salary: "$210k" } });
  check("discovery_add ok", !r.err);
  r = await call("discovery_update", { company: "Initech", position: "EM", fields: { Status: "Researching" } });
  check("discovery_update ok", !r.err);
  r = await call("promote_to_tracker", { company: "Globex", position: "EM Payments" });
  check("promote_to_tracker ok", !r.err);
  r = await call("get_job", { company: "Globex" });
  check("promoted row in tracker (Status Applied)", r.j.Status === "Applied");
  r = await call("promote_to_tracker", { company: "Globex", position: "EM Payments" });
  check("promote refuses duplicate", r.err);
  r = await call("discovery_delete", { company: "Initech", position: "EM" });
  check("discovery_delete ok", !r.err);
  r = await call("discovery_list", {});
  check("discovery_list count = 1 after churn", r.j.count === 1);

  console.log("\n# Interview prep (ambiguity)");
  r = await call("read_interview_prep", { company: "Acme" });
  check("read ambiguous Acme lists both", /multiple sections/i.test(r.text));
  r = await call("read_interview_prep", { company: "Acme — Role A" });
  check("read exact heading returns one section", /Why A\?/.test(r.text) && !/Why B\?/.test(r.text));
  r = await call("append_interview_prep", { company: "Acme", content: "**Q: X?**\n\nA: Y." });
  check("append ambiguous Acme refused", r.err);
  r = await call("append_interview_prep", { company: "Acme — Role A", content: "**Q: New?**\n\nA: Yes." });
  check("append to exact heading ok", !r.err && r.j.created === false);
  r = await call("append_interview_prep", { company: "Brand New Co", content: "**Q: Hi?**\n\nA: Ok." });
  check("append new section ok", !r.err && r.j.created === true);

  console.log("\n# Documents");
  // Payloads must start with real PDF magic — save_document rejects content that
  // doesn't match the extension (catches truncated base64).
  const pdfB64 = (body) => Buffer.from(`%PDF-1.4\n${body}`).toString("base64");
  r = await call("save_document", { filename: "Test.pdf", content_base64: pdfB64("one") });
  check("save_document ok", !r.err && r.j.via === "content_base64");
  r = await call("save_document", { filename: "Test.pdf", content_base64: pdfB64("two") });
  check("save refuses overwrite", r.err && /already exists/i.test(r.text));
  r = await call("save_document", { filename: "Test.pdf", content_base64: pdfB64("three"), overwrite: true });
  check("save overwrite makes backup", !r.err && r.j.backup);
  r = await call("save_document", { filename: "../evil.pdf", content_base64: pdfB64("x") });
  check("save blocks path traversal", r.err);

  // --- new: source_path is the preferred, token-free path ---
  const srcPdf = path.join(dir, "generated-resume.pdf");
  fs.writeFileSync(srcPdf, Buffer.from("%PDF-1.4\n" + "A".repeat(60_000)));
  r = await call("save_document", { filename: "FromPath.pdf", source_path: srcPdf });
  check(
    "save via source_path copies file",
    !r.err && r.j.via === "source_path" && r.j.bytes > 60_000 &&
      fs.existsSync(path.join(resumesDir, "FromPath.pdf"))
  );
  r = await call("save_document", { filename: "FromPath.pdf", source_path: srcPdf });
  check("source_path respects overwrite guard", r.err && /already exists/i.test(r.text));
  r = await call("save_document", { filename: "Missing.pdf", source_path: path.join(dir, "nope.pdf") });
  check("source_path missing file errors", r.err && /not found/i.test(r.text));
  r = await call("save_document", { filename: "Rel.pdf", source_path: "relative/path.pdf" });
  check("source_path must be absolute", r.err && /absolute/i.test(r.text));

  // --- new: guardrails that turn the old 5-minute hang into a fast error ---
  r = await call("save_document", {
    filename: "Huge.pdf",
    content_base64: Buffer.from("%PDF-1.4\n" + "A".repeat(40_000)).toString("base64"),
  });
  check("oversized base64 rejected fast", r.err && /source_path/i.test(r.text));
  r = await call("save_document", { filename: "Bad.pdf", content_base64: Buffer.from("not a pdf").toString("base64") });
  check("corrupt/mismatched bytes rejected", r.err && /doesn't look like/i.test(r.text));
  r = await call("save_document", { filename: "Neither.pdf" });
  check("neither source given errors", r.err && /source_path/i.test(r.text));
  r = await call("save_document", { filename: "Both.pdf", source_path: srcPdf, content_base64: pdfB64("x") });
  check("both sources given errors", r.err && /not both/i.test(r.text));
  r = await call("delete_document", { filename: "FromPath.pdf" });
  check("cleanup FromPath.pdf", !r.err);
  r = await call("list_documents", {});
  check("list_documents sees saved files", r.j.count >= 2 && r.j.documents.some((d) => d.name === "Test.pdf"));
  r = await call("list_documents", { filter: "amazon" });
  check("list_documents filter works", r.j.documents.every((d) => d.name.toLowerCase().includes("amazon")));
  r = await call("delete_document", { filename: "Test.pdf" });
  check("delete_document ok + backup", !r.err && r.j.backup);
  r = await call("delete_document", { filename: "Test.pdf" });
  check("delete missing refused", r.err);

  console.log("\n# Print error paths (no physical print)");
  r = await call("print_document", { filename: "nope.pdf" });
  check("print missing file refused", r.err);
  await call("save_document", { filename: "notes.txt", content_base64: Buffer.from("hi").toString("base64") });
  r = await call("print_document", { filename: "notes.txt" });
  check("print unsupported type refused", r.err);
  r = await call("print_document", { filename: resumeName, printer_name: "No Such Printer 9000" });
  // Message shape depends on the host: a machine with enumerable printers (the
  // normal case, including Windows CI runners' built-in "Microsoft Print to
  // PDF") lists them; a host with no print system at all (e.g. a Linux CI
  // runner without CUPS installed) reports that instead. Either is a correct,
  // non-crashing refusal — assert on that, not on one specific host's wording.
  check(
    "print unknown printer refused (with printer list, or a clear no-print-system error)",
    r.err && /available printers|print system|printer/i.test(r.j.message || ""),
    r.j.message
  );

  console.log("\n# generate_resume (auto-fallback to docx when no PDF converter)");
  // Default format is "pdf", but the tool must never hard-fail just because
  // no converter is installed: it should transparently fall back to .docx
  // and say so. Assert success either way, then check the format actually
  // produced matches what landed on disk.
  r = await call("generate_resume", {
    company: "Hooli",
    position: "Engineering Manager",
    summary: "Tailored summary for the target role.",
    key_qualifications: ["Relevant point one.", "Relevant point two."],
  });
  check("generate_resume succeeds regardless of PDF converter availability", !r.err, r.text);
  if (!r.err) {
    const ext = r.j.format === "docx" ? "docx" : "pdf";
    const f = path.join(resumesDir, `Dale-Magrath-Resume-Hooli-Engineering-Manager.${ext}`);
    const okFile =
      fs.existsSync(f) &&
      (ext === "docx" || fs.readFileSync(f).subarray(0, 5).toString("latin1") === "%PDF-");
    check(`generate_resume produced a valid .${ext} (format=${r.j.format})`, okFile, r.j.path);
    check("generate_resume echoes company/position", r.j.company === "Hooli" && r.j.position === "Engineering Manager");
    check(
      "generate_resume explains the fallback when a PDF converter was unavailable",
      r.j.format === "pdf" ? r.j.note === undefined : typeof r.j.note === "string" && /docx/i.test(r.j.note)
    );
    // Hooli is neither tracked nor a lead -> suggest add_job.
    check("nextStep suggests add_job for a new company", r.j.nextStep?.status === "new" && r.j.nextStep?.suggestedCall?.tool === "add_job");
    // Amazon is already in the tracker -> suggest update_job.
    const ra = await call("generate_resume", { company: "Amazon", position: "Senior Manager", summary: "s", key_qualifications: ["k"], overwrite: true });
    check("nextStep suggests update_job when already tracked", ra.j.nextStep?.status === "already_tracked" && ra.j.nextStep?.suggestedCall?.tool === "update_job");
    // A discovery-only lead (not yet in the tracker) -> suggest promote_to_tracker.
    await call("discovery_add", { company: "Contoso", position: "Eng Manager" });
    const rs = await call("generate_resume", { company: "Contoso", position: "Eng Manager", summary: "s", key_qualifications: ["k"], overwrite: true });
    check("nextStep suggests promote_to_tracker for a lead", rs.j.nextStep?.status === "in_discovery" && rs.j.nextStep?.suggestedCall?.tool === "promote_to_tracker");
  }
  // Explicitly requesting docx must always succeed, with no conversion attempt or note.
  const rDocx = await call("generate_resume", { company: "Wonka", position: "COO", summary: "s", key_qualifications: ["k"], format: "docx" });
  check("generate_resume with format:docx succeeds and skips conversion", !rDocx.err && rDocx.j.format === "docx" && rDocx.j.note === undefined, rDocx.text);
  // Missing key_qualifications is a schema error regardless of format/deps.
  r = await call("generate_resume", { company: "X", position: "Y", summary: "s" });
  check("generate_resume requires key_qualifications", r.err);

  console.log("\n# read_document");
  r = await call("read_document", { filename: "no-such.pdf" });
  check("read_document missing file errors", r.err);
  r = await call("read_document", { filename: "../evil.pdf" });
  check("read_document blocks path traversal", r.err);
  r = await call("read_document", { filename: "notes.txt" });
  check("read_document unsupported type errors + lists types", r.err && /\.pdf|\.docx/i.test(r.j.message || ""));
  // Read back whichever format generate_resume actually produced above
  // (.pdf if a converter was found, .docx if it fell back).
  const genPdf = path.join(resumesDir, "Dale-Magrath-Resume-Hooli-Engineering-Manager.pdf");
  const genDocx = path.join(resumesDir, "Dale-Magrath-Resume-Hooli-Engineering-Manager.docx");
  const genName = fs.existsSync(genPdf) ? "Dale-Magrath-Resume-Hooli-Engineering-Manager.pdf"
    : fs.existsSync(genDocx) ? "Dale-Magrath-Resume-Hooli-Engineering-Manager.docx"
    : null;
  if (genName) {
    r = await call("read_document", { filename: genName });
    check("read_document extracts readable text from the generated resume", !r.err && typeof r.j.text === "string" && /Test Candidate/.test(r.j.text));
    check("read_document reports words", !r.err && r.j.words > 0);
  } else {
    check("read_document generated-resume check skipped (no file found)", false, "neither .pdf nor .docx was produced above");
  }

  console.log("\n# Resume master field tools (scoped to resume_master.json)");
  r = await call("list_resume_master_structure", {});
  check("list_structure shows experience as array", r.j.structure?.fields?.experience?.type === "array");
  check("list_structure hides full content (name is type+length only)", r.j.structure?.fields?.name?.type === "string" && r.j.structure.fields.name.value === undefined);
  r = await call("get_resume_master_field", { path: "name" });
  check("get_field reads a scalar", r.j.value === "Test Candidate");
  r = await call("get_resume_master_field", { path: "experience[0].title" });
  check("get_field reads a nested path", r.j.value === "Senior Manager");
  r = await call("get_resume_master_field", { path: "nope.here" });
  check("get_field bad path errors", r.err);
  r = await call("update_resume_master_field", { path: "contact", value: "new@example.com • Toronto" });
  check("update_field ok + reports from/to", !r.err && r.j.from && r.j.to === "new@example.com • Toronto");
  r = await call("get_resume_master_field", { path: "contact" });
  check("update_field persisted", r.j.value === "new@example.com • Toronto");
  r = await call("update_resume_master_field", { path: "skills", value: ["X", "Y"] });
  check("update_field array->array ok", !r.err && Array.isArray(r.j.to));
  r = await call("update_resume_master_field", { path: "skills", value: "oops" });
  check("update_field rejects array->scalar type mismatch", r.err && /type mismatch/i.test(r.j.message || ""));
  r = await call("update_resume_master_field", { path: "brandNewKey", value: 1 });
  check("update_field refuses to create new top-level key", r.err);
  r = await call("update_resume_master_field", { path: "experience[0].title", value: "VP Engineering" });
  check("update_field edits nested path", !r.err);

  console.log("\n# read_text_file (scoped to Job Tracking root)");
  fs.mkdirSync(path.join(dir, "notes"), { recursive: true });
  fs.writeFileSync(path.join(dir, "notes", "plan.txt"), "my scratch plan");
  r = await call("read_text_file", { filename: "Interview_Prep_QA.md" });
  check("read_text_file reads a root .md", !r.err && /Acme/.test(r.j.content || ""));
  r = await call("read_text_file", { filename: "resume_master.json" });
  check("read_text_file reads .json", !r.err && /Test Candidate/.test(r.j.content || ""));
  r = await call("read_text_file", { filename: "notes/plan.txt" });
  check("read_text_file reads a subfolder path", !r.err && r.j.content === "my scratch plan");
  r = await call("read_text_file", { filename: "nope.md" });
  check("read_text_file missing lists existing files", r.err && /Interview_Prep_QA\.md/.test(r.j.message || ""));
  r = await call("read_text_file", { filename: "../escape.txt" });
  check("read_text_file blocks traversal outside root", r.err && /outside/i.test(r.j.message || ""));
  r = await call("read_text_file", { filename: "Job_Tracking.xlsx" });
  check("read_text_file refuses non-text extension", r.err);
  r = await call("read_text_file", { filename: "Resumes/x.pdf" });
  check("read_text_file points PDFs to read_document", r.err && /read_document/i.test(r.j.message || ""));

  console.log("\n# Gmail tools (no real inbox — creds point at missing files)");
  // send_gmail_email must refuse without confirm, before any auth/network.
  r = await call("send_gmail_email", { to: "x@example.com", subject: "s", body: "b" });
  check("send_gmail_email refuses without confirm:true", r.err && /confirm/i.test(r.j.message || ""));
  // Read-only tools fail gracefully (authorize-first), not with a crash.
  r = await call("search_gmail_for_job", { company: "Acme" });
  check("search_gmail_for_job errors gracefully when unauthorized", r.err && /auth|credential|gmail/i.test(r.j.message || ""));

  console.log("\n# Search criteria (network-free)");
  r = await call("get_search_criteria", {});
  check("get_search_criteria fresh install is incomplete", r.j.isComplete === false && Array.isArray(r.j.promptsNeeded) && r.j.promptsNeeded.length > 0);
  r = await call("get_broad_search_queries", {});
  check("get_broad_search_queries refuses before criteria set", r.err && /search criteria/i.test(r.j.message || ""));
  r = await call("run_job_sweep", {});
  check("run_job_sweep refuses before criteria set (no network hit)", r.err && /search criteria/i.test(r.j.message || ""));
  r = await call("update_search_criteria", {
    job_titles: ["Engineering Manager"],
    work_style: ["remote"],
    city: "Toronto",
    country: "Canada",
    min_salary: 150000,
    currency: "CAD",
  });
  check("update_search_criteria saves and reports complete", !r.err && r.j.isComplete === true);
  r = await call("get_search_criteria", {});
  check("get_search_criteria now complete", r.j.isComplete === true && r.j.criteria.jobTitles.includes("Engineering Manager"));
  r = await call("get_broad_search_queries", {});
  check("get_broad_search_queries succeeds once criteria complete", !r.err && Array.isArray(r.j.queries) && r.j.queries.length > 0);

  console.log("\n# discovery_sync (housekeeping + append, no network)");
  r = await call("discovery_sync", { new_rows: [] });
  check("discovery_sync housekeeping-only ok", !r.err && Array.isArray(r.j.appended) && typeof r.j.total_rows === "number");
  r = await call("discovery_sync", {
    new_rows: [{ company: "Netflix", position: "Director", job_link: "https://netflix.example/jobs/9" }],
  });
  check(
    "discovery_sync skips a candidate already in the tracker",
    !r.err && r.j.appended.length === 0 && r.j.skipped_duplicates.some((s) => /Netflix/.test(s) && /tracker/i.test(s))
  );
  r = await call("discovery_sync", {
    new_rows: [{ company: "Umbrella Corp", position: "VP Engineering", job_link: "https://umbrella.example/jobs/1" }],
  });
  check("discovery_sync appends a genuinely new candidate", !r.err && r.j.appended.some((s) => /Umbrella Corp/.test(s)));
  r = await call("discovery_sync", {
    new_rows: [{ company: "Umbrella Corp", position: "VP Engineering", job_link: "https://umbrella.example/jobs/1" }],
  });
  check("discovery_sync dedupes a repeat of the same candidate", !r.err && r.j.appended.length === 0 && r.j.skipped_duplicates.length === 1);

  console.log("\n# init_job_tracker_files / create_resume_master (idempotent guards)");
  r = await call("init_job_tracker_files", {});
  check(
    "init_job_tracker_files never touches existing tracker/discovery",
    !r.err && r.j.folderCreated === false && /already existed/i.test(r.j.tracker) && /already existed/i.test(r.j.discovery)
  );
  r = await call("create_resume_master", {
    name: "X", contact: "x@example.com", default_summary: "s",
    default_key_qualifications: ["k"], experience: [{ org: "O", title: "T", dates: "D", description: "d" }],
    education: ["E"], skills: ["S"],
  });
  check("create_resume_master refuses when one already exists", r.err && /already exists/i.test(r.j.message || ""));

  console.log("\n# check_setup (read-only readiness report)");
  r = await call("check_setup", {});
  check("check_setup reports existing files", !r.err && r.j.tracker.exists === true && r.j.discovery.exists === true && r.j.resumeMaster.exists === true);
  check("check_setup reflects completed search criteria", r.j.searchCriteria.isComplete === true);
  check("check_setup reports Gmail unauthorized (creds point nowhere)", r.j.gmail.authorized === false);
  check("check_setup reports a boolean pdfConverter.available regardless of host", typeof r.j.pdfConverter.available === "boolean");

  console.log("\n# Regression: repeated tracker writes must not balloon file size (cellStyles/theme bug)");
  const sizeBeforeRepeat = fs.statSync(trackerFile).size;
  for (let i = 0; i < 5; i++) {
    await call("update_job_status", { company: "Netflix", position: "Director", status: i % 2 === 0 ? "Interviewing" : "Applied" });
  }
  const sizeAfterRepeat = fs.statSync(trackerFile).size;
  check(
    "5 repeated writes keep tracker file size stable (no theme-doubling regression)",
    sizeAfterRepeat < sizeBeforeRepeat * 3 && sizeAfterRepeat < 500_000,
    `before=${sizeBeforeRepeat}B after=${sizeAfterRepeat}B`
  );

  console.log("\n# Backups rotate into .backups folder");
  const bdir = path.join(dir, ".backups");
  check(
    "tracker .backups created",
    fs.existsSync(bdir) && fs.readdirSync(bdir).some((f) => f.startsWith("Job_Tracking.xlsx."))
  );

  // Close stdin rather than force-killing: the server exits cleanly on its
  // own once stdin ends, which (unlike proc.kill() on Windows, where a forced
  // TerminateProcess skips exit handlers entirely) lets it flush its
  // NODE_V8_COVERAGE data — needed for `npm run coverage` to see anything
  // beyond this file itself. Falls back to a hard kill if it doesn't exit
  // promptly, so the suite never hangs.
  proc.stdin.end();
  await new Promise((resolve) => {
    proc.once("exit", resolve);
    setTimeout(resolve, 3000);
  });
  if (proc.exitCode === null) proc.kill();
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch {}

  console.log(`\n${failed === 0 ? "ALL PASSED" : "FAILURES"}: ${passed} passed, ${failed} failed`);
  if (failed) { console.log("Failed:", fails.join(", ")); process.exit(1); }
}

main().catch((e) => { console.error("Test harness error:", e); process.exit(1); });
