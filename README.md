# Job Search Assistant — MCP Server

A local [Model Context Protocol](https://modelcontextprotocol.io) server that lets
Claude read and update your job-search spreadsheets through natural conversation —
"what's still awaiting a response after 2 weeks?", "add this posting to my
tracker", "mark the Amazon role as interviewing", "add a Salary Range column and
set it to 180-210k for the Amazon role", "what remote leads did I find this week?",
"promote the Acme Corp lead to my tracker".

It manages a funnel — **discover → apply → track** — plus an interview-prep note:

| File | Role | Env var |
|------|------|---------|
| `Job_Tracking.xlsx` | Applications you've submitted (the **tracker**) | `JOB_TRACKER_FILE` |
| `Job_Search_Discovery.xlsx` | Leads you've found but not yet applied to (**discovery**) | `JOB_DISCOVERY_FILE` |
| `Interview_Prep_QA.md` | Rolling interview Q&A reference (**prep**) | `JOB_INTERVIEW_PREP_FILE` |
| `Resumes\` | Folder for generated resumes / cover letters, saved and printed from | `JOB_RESUMES_DIR` |
| `resume_master.json` | Stable resume facts that `generate_resume` tailors from | `JOB_RESUME_MASTER_FILE` |

`promote_to_tracker` turns a discovery lead into a tracked application; `save_document`
and `print_document` write generated docs to the `Resumes\` folder and print them. An
optional **Gmail integration** (five tools) searches your inbox for recruiter mail,
scans for updates on tracked roles, and drafts/sends replies — see *Gmail integration*.

- **Runtime:** Node.js + TypeScript — **no Python or other external interpreter required**
- **Transport:** stdio (local process — no networking, no auth)
- **Data source:** your `.xlsx` files, read/written with [SheetJS (`xlsx`)](https://sheetjs.com); styling (bold header, freeze/auto-filter, hyperlinks) via [`exceljs`](https://www.npmjs.com/package/exceljs)
- **Schema:** dynamic — driven by each sheet's header row (custom columns supported)
- **Resumes:** `.docx` built in-process with the [`docx`](https://www.npmjs.com/package/docx) package; `.pdf`/`.docx` text extraction via [`pdf-parse`](https://www.npmjs.com/package/pdf-parse) + [`mammoth`](https://www.npmjs.com/package/mammoth)
- **Board sweep:** `run_job_sweep` probes ATS job boards directly over HTTP (native `fetch`), in-process — no external script
- **Printing:** silent PDF printing — [`pdf-to-printer`](https://www.npmjs.com/package/pdf-to-printer) (bundles SumatraPDF) on Windows, CUPS's `lp` on macOS/Linux; `.docx` is converted to PDF first using LibreOffice (any OS) or Microsoft Word (Windows only) if installed (optional — see *Safety & behavior notes*)
- **Client:** Claude Desktop or Claude Code (see config below)

## Build

```bash
cd mcp-job-tracker
npm install
npm run build
```

This produces `dist/index.js`. `npm install` pulls everything this server needs —
`pdf-to-printer` (bundles SumatraPDF for silent Windows printing, no separate
install; macOS/Linux print through CUPS's `lp` instead, no npm package needed),
`exceljs`, `docx`, `pdf-parse`, `mammoth` — all pure JS, nothing to compile and no
Python. The one optional *external* dependency is **LibreOffice** (any OS) or
**Microsoft Word** (Windows only), used only to convert a generated resume's
`.docx` to `.pdf`; if none is installed, `generate_resume` still produces the
`.docx` and says so. The optional
Gmail tools need a one-time `npm run gmail:auth` (see *Gmail integration*).

## Configure Claude Desktop

Open Claude Desktop's config file (Windows):

```
%APPDATA%\Claude\claude_desktop_config.json
```

Add a `job-tracker` server entry (create the file / `mcpServers` object if needed):

```json
{
  "mcpServers": {
    "job-tracker": {
      "command": "node",
      "args": ["C:\\Users\\developer\\mcp-job-tracker\\dist\\index.js"],
      "env": {
        "JOB_TRACKER_FILE": "C:\\Users\\developer\\Documents\\Job_Tracking.xlsx",
        "JOB_DISCOVERY_FILE": "C:\\Users\\developer\\Documents\\Job_Search_Discovery.xlsx",
        "JOB_INTERVIEW_PREP_FILE": "C:\\Users\\developer\\Documents\\Interview_Prep_QA.md",
        "JOB_RESUMES_DIR": "C:\\Users\\developer\\Documents\\Resumes"
      }
    }
  }
}
```

Restart Claude Desktop. You should see the `job-tracker` tools appear. If
`JOB_TRACKER_FILE` is omitted, the server defaults to
`<home>\Documents\Job_Tracking.xlsx`. If `JOB_DISCOVERY_FILE`,
`JOB_INTERVIEW_PREP_FILE`, or `JOB_RESUMES_DIR` are omitted, they default to
`Job_Search_Discovery.xlsx`, `Interview_Prep_QA.md`, and a `Resumes\` folder
sitting **next to** the tracker file. **Claude Code** uses the same entry under
`mcpServers` in `~/.claude.json`.

## Moving to another machine

The server is portable (plain Node stdio — Windows or macOS). To set it up on a
new machine:

1. Copy this project folder over (you can skip `node_modules` and `dist`).
2. Copy your `Job_Tracking.xlsx` to the new machine — or, for a genuinely fresh
   start with no existing spreadsheet, skip this step and call the
   `init_job_tracker_files` tool once the server is running; it creates the
   folder and both blank, formatted workbooks for you.
3. Make sure Node.js 18+ is installed (`node --version`). Nothing else — no
   Python, no other interpreter — is required for the server itself.
4. Run the installer, pointing it at your spreadsheet:

   ```bash
   node setup.mjs "C:\\path\\to\\Job_Tracking.xlsx"     # Windows
   node setup.mjs "/Users/you/Documents/Job_Tracking.xlsx"  # macOS
   ```

   It installs deps, builds, backs up the existing Claude Desktop config, and
   wires in the `job-tracker` entry with paths correct for that machine. Omit the
   path to default to `<home>/Documents/Job_Tracking.xlsx`.
   Flags: `--no-build` (reuse a copied `dist/`), `--bare-node` (use `node` from
   PATH instead of an absolute node path).

5. Fully quit and restart Claude Desktop.

## Tools

### Setup

| Tool | What it does |
|------|--------------|
| `init_job_tracker_files` | Bootstrap a fresh machine/profile: creates the Job Tracking root folder if missing, and creates `Job_Tracking.xlsx` and/or `Job_Search_Discovery.xlsx` from a blank, formatted template (header row only) for whichever is missing. Never touches a file that already exists. |

| Tool | What it does |
|------|--------------|
| `list_jobs` | List all rows. Optional filters: `status` (case-insensitive exact), `company` (substring), and a `applied_from` / `applied_to` date range (YYYY-MM-DD). |
| `get_job` | Full details for one application, matched by `company` (+ `position` if the company has several rows). Includes `daysSinceApplied` and `resumeFiles` (matching files in the Resumes folder, so you can see which resume you used). |
| `add_job` | Append a new row. `company` + `position` required; everything else optional. `status` defaults to **Awaiting Response**, `date_applied` defaults to **today**. Set custom columns via `extra_fields`. |
| `update_job_status` | Update `status` (and optionally `notes`) for a row matched by `company` + `position`. Refuses to touch a non-existent or ambiguous match. |
| `update_job` | Edit **any** field(s) of a row matched by `company` + `position` — location, resume version, dates, notes, and custom columns (via `extra_fields`). Rename via `new_company` / `new_position`; pass an empty string to clear a field. Only supplied fields change; reports exactly what changed. |
| `delete_job` | Permanently remove a row matched by `company` + `position`, shifting rows below it up (like Excel's Delete Row). Refuses missing/ambiguous matches, backs up first, and returns the full deleted row so it can be re-added. |
| `get_stale_jobs` | Rows still *Awaiting Response* or *Applied* whose Date Applied is more than `days` ago (default **14**), computed against today. |
| `get_due_followups` | Rows whose **Next Follow-Up** date is due — on/before today, or within the next `days` days. Skips closed rows (Declined/Withdrawn/Closed); sorted most-overdue first (`daysUntil` is negative when overdue). |
| `search_jobs` | Case-insensitive substring search across **Company**, **Position**, and **Notes**. |
| `draft_followup_message` | *(stretch)* Drafts a short, professional follow-up email for a role, referencing time since applying and the resume version used. **Draft text only — it does not send anything.** |
| `add_column` | Add a new custom column at the end of the sheet. Optional `default_value` fills every existing row. Rejects duplicate names. |
| `modify_column` | Rename a custom column (`new_name`) and/or set one value across every row (`fill_value`; empty string clears it). Refuses on the base columns. |
| `delete_column` | Remove a custom column, shifting the columns to its right left (like Excel's Delete Column). Refuses on the base columns. |

### Discovery-sheet tools (`Job_Search_Discovery.xlsx`)

| Tool | What it does |
|------|--------------|
| `discovery_list` | List leads. Optional filters: `company` (substring), `status` (exact), `query` (substring across all columns). |
| `discovery_add` | Append a lead. `company` + `position` required; all other columns via `fields`. `Date Found` defaults to today, `Status` to **Open**. |
| `discovery_update` | Edit a lead matched by `company` + `position`. Rename via `new_company` / `new_position`; set columns via `fields` (empty string clears). |
| `discovery_delete` | Remove a lead matched by `company` + `position`, returning the deleted row so it can be re-added. |
| `promote_to_tracker` | Copy a lead into the tracker as an application — maps Company/Position/Job Link/Location, sets Status **Applied** and Date Applied **today** (both overridable), and folds Salary + Match Assessment into Notes. Refuses to create a duplicate tracker row. Optional `remove_from_discovery` moves it instead of copying. |
| `discovery_sync` | Housekeeping + append in one call: drops discovery rows whose Company+Position now matches a tracker application, flags anything 60+ days old **Stale (60+ days)**, then appends `new_rows` (refusing duplicates by Company+Position, Job Link, or an existing tracker application). Pass an empty array to run housekeeping alone. Prefer this over repeated `discovery_add` calls when processing a sweep's results. |

### Search criteria & board sweep

| Tool | What it does |
|------|--------------|
| `get_search_criteria` | Return the saved search criteria (`search_criteria.json`): target job titles, work style (remote/hybrid/onsite), city/country, minimum annual salary, and a free-text `notes` field. On a fresh install nothing is saved yet — the result's `isComplete` is `false` and `promptsNeeded` lists exactly what to ask the user before calling `update_search_criteria`. |
| `update_search_criteria` | Create or update the saved criteria. Every field is optional — only what you pass changes. Backs up the previous file first. |
| `run_job_sweep` | Probe ~95 Greenhouse/Ashby/Lever/SmartRecruiters/Workday boards concurrently (in-process, plain HTTP — no external script), filter to engineering-leadership titles whose location plausibly matches the saved criteria, fetch salary/posted-date detail, HTTP-check links, and dedupe against both spreadsheets. Returns only new candidates. Refuses to run until search criteria are complete. Optional `quick: true` skips boards the last run recorded as unreachable. |

### Interview-prep tools (`Interview_Prep_QA.md`)

| Tool | What it does |
|------|--------------|
| `read_interview_prep` | Return the prep file as text, or just one company's section with the optional `company` filter. Friendly message (not an error) if the file or section doesn't exist. |
| `append_interview_prep` | Append markdown under a `## {company}` section — creating the file (with a top-level heading) and/or the section as needed, or adding to an existing section without duplicating its heading. Backs up first. |

### Document tools (`Resumes\`)

| Tool | What it does |
|------|--------------|
| `generate_resume` | **Generate a tailored resume PDF (or docx) host-side and save it to `Resumes\` — no base64, works from Claude Desktop.** Stable facts (employers, dates, education, skills) come from `resume_master.json`; the model supplies only the per-posting `summary` + `key_qualifications` bullets for a `company` + `position`. The `.docx` is built in-process (the `docx` package); PDF output converts that via LibreOffice or Word if either is installed. See the workflow below. |
| `save_document` | Save a `.docx`/`.pdf` into the `Resumes\` folder (created if missing). Takes **`source_path`** (preferred — an absolute path to a file already on disk, copied instantly) *or* `content_base64` (fallback, small files only — capped at 20 000 chars ≈ 15 KB). Validates the leading bytes match the extension. Refuses to overwrite unless `overwrite: true`, backing the old file up first. Returns the saved path and size. |
| `print_document` | Silently print a saved file (name in `Resumes\`, or a full path) to the default or a named printer. PDFs print directly; `.docx` is converted to PDF first (cached next to the docx). Lists available printers if a bad name is given. Confirms the job was **sent**, not physically finished. |
| `read_document` | Extract the plain text of a saved `.pdf`/`.docx` in `Resumes\` (same scoping as `delete_document`) so it can be read/analyzed in conversation. Read-only; returns text + metadata (type, pages, words, chars). Very long text is capped with an explicit `truncated` note — never silently. |
| `read_text_file` | Read any plain-text file (`.md`, `.txt`, `.json`, `.csv`, …) anywhere under the **Job Tracking root** — handoff notes, scratch files, etc. — by filename or relative path. Read-only and strictly scoped: `..`/absolute paths that escape the root are refused; a missing file lists the directory's contents. Non-text extensions are refused (PDFs/docx → `read_document`). Truncates past ~1 MB with a note; refuses past 5 MB. |
| `list_documents` | List saved resumes / cover letters in the `Resumes\` folder (name, size, modified date). Optional `filter` substring on the name; newest first. |
| `delete_document` | Delete a document from the `Resumes\` folder, backing it up (rotating `.backups`) first. Refuses a missing file or a path outside the folder. |
| `list_resume_master_structure` | Show the *shape* of `resume_master.json` (keys, array lengths, scalar types — not full values) to find a path. Optional `path` scopes it. |
| `get_resume_master_field` | Read one field from `resume_master.json` by dot/bracket path, e.g. `skills`, `projects[0].description`, `experience[2].title`. |
| `update_resume_master_field` | Update one **existing** field in `resume_master.json` by path (backs up first, reports old → new). Won't create new keys or change a field's kind (array/object/scalar). Scoped to that one file. |

### Gmail tools (optional — require one-time auth)

| Tool | What it does |
|------|--------------|
| `search_gmail_for_job` | Search your Gmail for job-related mail by `company` (+ optional `position`, free-text `query`, and `limit`, default 10 / max 50). **Read-only.** |
| `read_gmail_message` | Read one message's full body by its `id` (e.g. from a `search_gmail_for_job` result). **Read-only.** |
| `scan_job_updates` | Scan recent Gmail for updates relevant to your tracked applications — optional `company` filter, `days` window (default 30), per-company `limit` (default 3). **Read-only.** |
| `draft_gmail_reply` | Save a Gmail **draft** — reply within a `thread_id`, or a new message to `to`; `body` required, optional `subject`/`cc`. **Never sends.** |
| `send_gmail_email` | **Sends immediately and cannot be recalled.** Refuses unless `confirm: true`, and only after the user has approved the exact to/subject/body. |

Until Gmail is authorized (see **Gmail integration** below), all five return a clear "authorize first" error.

## Spreadsheet schema

Both workbooks use a **dynamic** schema — the columns are driven by each sheet's
actual header row, and the read/write engine is schema-agnostic (a sheet is just
a spec: file path, required columns, which columns are dates, and the match keys).

### Tracker (`Job_Tracking.xlsx`)

The ten
**base columns** below must always be present (the tools match, filter, and
compute dates on them) and are protected from rename/delete:

```
Company | Position | Job Link | Location | Resume Version | Contact/Referral | Date Applied | Status | Next Follow-Up | Notes
```

Any **custom columns** you add beyond these are first-class:

- They appear in `list_jobs` / `get_job` output automatically.
- You set them per job with the `extra_fields` argument of `add_job` and
  `update_job`, e.g. `extra_fields: { "Salary Range": "180-210k" }` (an empty
  string clears the value).
- You add / rename / remove them with `add_column`, `modify_column`, and
  `delete_column`.

Custom columns are stored as text. The two base date columns (**Date Applied**,
**Next Follow-Up**) are the ones stored as Excel date serials.

### Managing columns

```jsonc
// add a column, pre-filled on every existing row
add_column      { "name": "Salary Range", "default_value": "TBD" }
// rename it and/or overwrite every row's value
modify_column   { "name": "Salary Range", "new_name": "Comp", "fill_value": "Confidential" }
// remove it (shifts later columns left, like Excel)
delete_column   { "name": "Comp" }
```

Conventional `Status` values: `Applied`, `Awaiting Response`,
`Informal - Referral Sent`, `Interviewing`, `Offer`, `Declined`, `Withdrawn`,
`Closed - No Longer Available`. (The field is free text — these are just the
values suggested to Claude.)

### Discovery (`Job_Search_Discovery.xlsx`)

The base columns, in order:

```
Date Found | Company | Position | Location (Remote/Hybrid) | Salary | Salary Confidence (Confirmed/Estimated) | Job Link | Source (Company Careers Page/LinkedIn/Aggregator) | Posted Date / Days Since Posted | Known Gap Flag | Match Assessment | Status
```

All discovery values are stored as **text** (including `Date Found`, kept as an
ISO `YYYY-MM-DD` string — no date-serial conversion). Rows are matched by
**Company + Position**. Set columns through the `fields` argument of
`discovery_add` / `discovery_update`; `Company` and `Position` have their own
parameters. `Status` on new leads defaults to **Open**.

The column-management tools (`add_column` etc.) operate on the **tracker**; the
discovery sheet is managed through its own row tools.

**Formatting:** SheetJS (used for the data read/write itself) doesn't
re-serialize rich cell styling, so every MCP discovery write is followed by a
best-effort **`exceljs` reformat pass** (`xlsxFormat.ts`) that (re)applies bold
header / frozen row / auto-filter / wrapped text / clickable Job Link cells.
Pure JS, nothing external required; if it ever fails for some reason the data
write still succeeds (the tool returns `formatted: false`).

### Interview prep (`Interview_Prep_QA.md`)

A plain-markdown reference — not a spreadsheet — organised as:

```markdown
# Interview Prep — Q&A Reference

## Company — Role
**Q: ...?**

A: ...

*Notes: ...*

---

## Next Company — Role
...
```

`append_interview_prep` matches an existing `## Company` section on the company
portion of the heading (the part before ` — Role`), so appending "Acme Corp" lands
under `## Acme Corp — Engineering Manager`. New sections are separated with `---`.
Keep to the `**Q: …**` / `A: …` / optional `*Notes: …*` pattern when adding content.

If a bare company name matches **more than one** section (e.g. `## Acme Corp — Billing`
and `## Acme Corp — Clinical`), both `read`/`append` refuse and list the matches — pass
the **full heading** to disambiguate. An exact full-title match always wins over a
bare prefix, so a specific heading is never ambiguous.

### Generating a tailored resume (`generate_resume`)

The recommended way to produce a resume — especially from **Claude Desktop**,
which can't write to arbitrary host paths. The model never emits the binary;
it emits *tailored text*, and the server renders the PDF host-side. The server's
MCP `instructions` tell connected clients to use this tool for resumes rather
than hand-rolling their own resume-generation script or moving bytes through
base64, so it's the single resume path (`resumeDocx.ts` is the one generator).

```jsonc
generate_resume {
  "company": "Acme Corp",
  "position": "Engineering Manager",
  "summary": "…tailored SUMMARY paragraph for this posting…",
  "key_qualifications": [
    "…bullet tying real experience to this role…",
    "…another alignment bullet…"
  ]
  // optional: skills[], include_projects, format ("pdf"|"docx"), filename, overwrite
}
```

- **Facts stay fixed.** Employers, titles, dates, education, certs, and the base
  skills live in `resume_master.json`; the model can't invent or drift them — it
  only supplies the `summary` and `key_qualifications` for this `company` +
  `position`. Edit your real facts in `resume_master.json` (a sibling of the
  tracker; override with `JOB_RESUME_MASTER_FILE`).
- **Rendering** merges master + tailoring → `.docx` via `resumeDocx.ts` (the
  `docx` package, in-process; the measured spec: Liberation Serif, 20 pt name,
  ruled 11 pt headings, 0.63" margins, US Letter) → PDF via LibreOffice or Word,
  whichever is installed. ~15 s per resume.
- **Output** lands in `Resumes\` (e.g. `Dale-Magrath-Resume-Acme-Corp-Engineering-Manager.pdf`),
  ready for `print_document`. No base64, no download-then-move.
- **Wired to tracking.** The result carries a `nextStep` with the exact follow-up
  call, chosen from the sheets: brand-new company → `add_job` (as Applied, Resume
  Version set to the file), an existing discovery lead → `promote_to_tracker`,
  already tracked → `update_job` to set the Resume Version. So *tailor → save →
  apply → track* is one Desktop flow — the model just offers the suggested step.

**Why this beats base64:** the tailored text is ~1 k tokens; the 55 KB PDF binary
would be ~22 k. The model does the writing; the host does the rendering.

#### Rendering notes (why it's built this way)

Hard-won details worth keeping — changing them tends to reintroduce old bugs:

- **Role/date lines are a 2-column borderless table, not a tab stop.** A tab with
  a right tab-stop spanning a **bold** title and *italic* dates in one paragraph
  triggered a LibreOffice PDF-export bug (ToUnicode CMap on subset fonts) that made
  the title extract as **garbled characters** in `pdftotext`, `pypdf`, and
  `pdfminer.six` — even though it looked fine on screen (and could trip ATS
  parsers). The table renders identically and extracts cleanly in all three. Keep
  it a table.
- **LibreOffice does the docx→PDF, not Word.** LibreOffice subsets fonts →
  ~55 KB with Liberation Serif, matching the existing resumes. Word barely subsets
  (embedded ~1.1 MB of Calibri → a 258 KB PDF), so `convertDocxToPdf` prefers
  LibreOffice and falls back to Word only if it's absent — and only on Windows;
  macOS/Linux have no COM automation to fall back to, so LibreOffice is the sole
  path there (checked at `/Applications/LibreOffice.app/...` on macOS, common
  Linux install paths, then bare `soffice` on PATH).
- **2-page fit is by layout, never by truncation.** A full experience history +
  6-bullet Key Qualifications + Projects + Education + Skills lands on 2 pages with
  **Honors & Awards off** (opt in with `include_honors`). If it ever runs long, the
  lever is tightening spacing (heading 12/5 pt, body 8 pt) — content is never
  silently trimmed to hit a page count.
- **base64 is a last resort.** The only resume ever saved through the base64 path
  came out **corrupt** (valid `%PDF`/`%%EOF` but a broken `startxref`). The
  leading-bytes check catches gross truncation, not subtle corruption — so prefer
  `generate_resume`, or `save_document` with `source_path`, for anything binary.

### Saving an already-generated document (`Resumes\`)

If a file already exists on disk (e.g. produced by the docx/pdf skill),
`save_document` puts it into `Resumes\`, and `print_document` sends it to a printer:

```jsonc
// Preferred: hand over a path to a file already written to disk.
save_document   { "filename": "Dale-Resume-Acme-Corp.docx", "source_path": "C:\\Temp\\out.docx" }
print_document  { "filename": "Dale-Resume-Acme-Corp.docx" }            // default printer
print_document  { "filename": "sample.pdf", "printer_name": "Office EPSON ET-3850 Series" }
```

> **Always prefer `source_path` for real documents.** `content_base64` makes the
> model emit the whole file as base64 token-by-token — roughly 3 characters per
> token for binary formats, so a 53 KB resume is ~22 000 output tokens (several
> minutes), and a 134 KB PDF exceeds the response limit and can never finish.
> `source_path` costs ~30 tokens at any file size and can't corrupt the bytes.
> Inline base64 is therefore capped at 20 000 characters and fails immediately
> with a pointer to `source_path`.

- **PDFs** print directly — via SumatraPDF (bundled with `pdf-to-printer`) on
  Windows, or CUPS's `lp` on macOS/Linux (present by default on macOS, and on
  any Linux with CUPS installed). No dialog, single copy, default settings
  either way.
- **`.docx`** is converted to a PDF next to it first — using LibreOffice
  (`soffice --headless --convert-to pdf`) if installed, otherwise (Windows
  only) Microsoft Word via COM. The PDF is cached and reused while it's newer
  than the docx. If no converter is available, printing a `.docx` returns a
  clear error suggesting you save it as a PDF instead.
- `print_document` confirms the job was **sent** to the printer — not that it
  physically finished. Base64 transport is meant for small files (resumes,
  well under 1 MB), not large documents.

## Gmail integration (optional)

The five Gmail tools let Claude search your inbox for recruiter mail, scan for
updates on tracked applications, and draft/send replies — all against your own
account. They're **off until you authorize** them; unauthorized calls return a
clear error telling you what to run.

**One-time setup** (full walkthrough in [`GMAIL_SETUP.md`](GMAIL_SETUP.md)):

1. Create a Google Cloud project, enable the Gmail API, and create a **Desktop
   app** OAuth client; download its JSON as `credentials.json` in the project
   root (next to `package.json`). This step is manual — it can't be scripted.
2. Run the one-time consent flow:
   ```bash
   npm run gmail:auth
   ```
   It opens Google's consent screen and saves the token to `gmail_token.json`.
   The server refreshes it automatically afterward.
3. Restart Claude Desktop so the tools load.

**Config / files** (both `.gitignore`d):

| File | Purpose | Env override |
|------|---------|--------------|
| `credentials.json` | OAuth client (manual download) | `JOB_GMAIL_CREDENTIALS_FILE` |
| `gmail_token.json` | Stored access/refresh token from consent | `JOB_GMAIL_TOKEN_FILE` |

Requested scopes are **read-only + compose + send** (`gmail.readonly`,
`gmail.compose`, `gmail.send`) — intentionally **no** mailbox modify/delete
access. `npm install` pulls the one dependency (`google-auth-library`).

**Safety:** `search_gmail_for_job` / `read_gmail_message` / `scan_job_updates`
are read-only; `draft_gmail_reply` only ever creates a draft; `send_gmail_email`
**sends immediately and can't be undone**, so it refuses without `confirm: true`
and should only be called after you've reviewed the exact to/subject/body.

## Safety & behavior notes

- **Backups (rotating):** every write first copies the file into a sibling
  `.backups\` folder with a timestamped name (`<file>.<YYYYMMDD-HHmmss-SSS>.bak`),
  keeping the **last 10** per file. This replaces the old single-slot `.bak`, so
  repeated writes (manual edits or the daily task) can't clobber your last-good
  copy. Applies to the sheets, the prep file, and `delete_document`;
  `promote_to_tracker` backs up each file it writes. (Nothing is backed up when a
  file is created fresh.)
- **Fresh reads:** the workbook is re-read on every call, so edits you make
  directly in Excel are always reflected.
- **Validation:** row-editing tools verify the row exists (and is unambiguous)
  before changing anything; bad dates are rejected up front. `promote_to_tracker`
  refuses to create a duplicate tracker row.
- **Dates** (tracker only) are stored as Excel date serials formatted `d-mmm-yy`.
  New rows are written the same way, so `2026-06-01` shows as `1-Jun-26` —
  matching the existing rows. Reads/writes use timezone-free calendar-date math.
  Discovery dates are plain ISO text.
- **Formatting preservation:** new rows copy the number format (and, where the
  writer supports it, the cell style) of the last existing row, and column widths
  are preserved — the way a human copy-pasting a row would leave things.
  > Note: the open-source SheetJS build faithfully preserves **number formats**
  > and **column widths** (what this workbook uses). Rich per-cell styling
  > (custom fonts/fills/borders) is not re-serialized by the community build; if
  > you later add heavy cell styling, preserving it would require SheetJS Pro.
- **File locked:** if a file is open in another program (Excel for the sheets,
  Word/an editor for the prep `.md`), writes fail with a clear "please close it"
  message rather than a cryptic error.

## Testing

```bash
npm test
```

`test-client.mjs` builds the server, then spins it up over stdio against
**self-contained fixtures in a temp folder** and asserts on every tool (66
checks across the tracker, discovery, promote, interview-prep ambiguity,
document, and backup paths). Your real files are never touched — the suite
creates its own throwaway tracker/discovery/prep/Resumes in the OS temp dir and
deletes them afterward. It exits non-zero if any check fails.

It does **not** send a physical print job — `print_document` is covered only via
its error paths (missing file, unsupported type, unknown printer).
