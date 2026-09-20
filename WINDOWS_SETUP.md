# Windows 11 setup (Claude Desktop)

A complete walkthrough for getting this server running on a Windows 11
machine that has nothing installed yet — no Node, no Claude Desktop, no
existing spreadsheet. If some of these are already installed, skip ahead.

Total time: ~10 minutes, plus however long `npm install` takes on your
connection.

## What you'll end up with

- Claude Desktop, with a `job-tracker` MCP server wired into its config
- This project's code, built, living wherever you put it (e.g.
  `C:\mcp-job-tracker`)
- A blank `Job_Tracking.xlsx` / `Job_Search_Discovery.xlsx` in your Documents
  folder (or an existing one you already have, copied into place)

## 1. Install Node.js

1. Go to https://nodejs.org/ and download the **LTS** installer (the button
   labeled "LTS", not "Current").
2. Run the installer, accepting the defaults — it's fine to leave every
   checkbox (including "Automatically install the necessary tools") at its
   default.
3. Open a **new** PowerShell window (Start menu → type `PowerShell` → Enter)
   and verify:

   ```powershell
   node --version
   npm --version
   ```

   Both should print a version number (Node 18 or later). If PowerShell says
   `node is not recognized`, close and reopen PowerShell — the installer
   updates PATH, but only windows opened afterward see it.

## 2. Get the project files

Pick one:

**Option A — Git (recommended if you have it, or don't mind installing it):**

```powershell
cd C:\
git clone https://github.com/DaleMagrath/mcp-job-tracker.git
cd mcp-job-tracker
```

Don't have Git? Get it from https://git-scm.com/download/win (defaults are
fine), then use the commands above from a new PowerShell window.

**Option B — Download a ZIP (no Git needed):**

1. Go to https://github.com/DaleMagrath/mcp-job-tracker
2. Click the green **Code** button → **Download ZIP**.
3. Right-click the downloaded ZIP → **Extract All...** → somewhere simple
   like `C:\mcp-job-tracker` (avoid extracting straight into `Downloads`).
4. In PowerShell:

   ```powershell
   cd C:\mcp-job-tracker
   ```

   (adjust if you extracted it somewhere else — if the ZIP nested itself in
   an extra `mcp-job-tracker-main` folder, `cd` into that instead)

## 3. Install Claude Desktop (if you don't have it)

Download and install from https://claude.ai/download. Open it once so it
creates its own config folder, then fully quit it (right-click its icon in
the system tray, bottom-right of the taskbar → **Quit** — closing the window
alone leaves it running in the background).

## 4. Install dependencies and build

From the project folder in PowerShell:

```powershell
npm install
npm run build
```

`npm install` needs internet access — one of its dependencies (the `xlsx`
library) is fetched directly from `cdn.sheetjs.com` rather than the usual npm
registry, so a restrictive corporate firewall/VPN can block just that one
package even if everything else installs fine. `npm run build` compiles the
TypeScript to `dist\index.js`; it should finish with no output (silence means
success — `tsc` only prints on error).

**If `npm install` or `npm run build` fails with something like "cannot be
loaded because running scripts is disabled on this system"** — that's
PowerShell's execution policy blocking npm's script shim, not an actual
problem with the project. Either:
- Run the same commands from **Command Prompt** (`cmd.exe`) instead of
  PowerShell, or
- Run PowerShell as Administrator once and execute:
  `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then retry from a
  normal PowerShell window.

## 5. Wire it into Claude Desktop

From the same PowerShell window, run the one-shot installer:

```powershell
node setup.mjs
```

This wires a `job-tracker` entry into Claude Desktop's config
(`%APPDATA%\Claude\claude_desktop_config.json`), pointing it at
`dist\index.js` and defaulting your data to
`C:\Users\<you>\Documents\Job_Tracking.xlsx`. It backs up any existing config
first (`claude_desktop_config.json.bak`) rather than overwriting it blindly.

**Already have a `Job_Tracking.xlsx`** from another machine or an earlier
setup? Pass its path instead of using the default:

```powershell
node setup.mjs "C:\Users\<you>\Documents\Job_Tracking.xlsx"
```

(if it's not already in place, copy it there first — `setup.mjs` just
records the path, it doesn't copy the file for you.)

You don't need a spreadsheet at all to proceed — an empty/missing one is
fine, see step 7.

## 6. Restart Claude Desktop and verify

Fully quit Claude Desktop (system tray → **Quit**, not just closing the
window) and reopen it. Start a new conversation and check the tools loaded —
either look for a small hammer/tools icon near the message box, or just ask:

> what job-tracker tools do you have available?

If nothing shows up, see *Troubleshooting* below.

## 7. First things to try

In a new conversation, ask Claude to run the setup check:

> can you check my job-tracker setup?

That calls `check_setup` — a read-only tool that reports exactly what
exists, what's still missing (no spreadsheet yet? no resume master? no PDF
converter?), and what to do about each one. Two likely first steps it'll
point you at:

- **No spreadsheet yet:** ask Claude to run `init_job_tracker_files` — it
  creates a blank, correctly-formatted `Job_Tracking.xlsx` and
  `Job_Search_Discovery.xlsx` in your Documents folder for you.
- **No resume master yet:** upload your current resume (PDF or DOCX) and ask
  Claude to set up `resume_master.json` from it — needed before
  `generate_resume` will work.

## 8. Optional: PDF resumes (LibreOffice or Microsoft Word)

`generate_resume` always produces a `.docx`. For it to also produce a `.pdf`,
install one of:

- **[LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)**
  — works out of the box once installed, no configuration needed.
- **Microsoft Word**, if you already have it — no separate download; the
  server automates it directly.

Neither is required — without one, `generate_resume` just saves the `.docx`
and says so in its result.

## 9. Optional: Gmail integration

Five tools (`search_gmail_for_job`, `scan_job_updates`, etc.) can read/draft/
send Gmail on your behalf. This needs a one-time Google Cloud OAuth setup —
see [GMAIL_SETUP.md](GMAIL_SETUP.md) for the full walkthrough. Skip this
entirely if you don't want Claude touching your inbox; every other tool works
fine without it.

## Troubleshooting

**Tools don't appear in Claude Desktop at all.**
- Confirm you fully quit and restarted the app (system tray → Quit, not just
  closing the window) — Claude Desktop only reads its config file at launch.
- Open `%APPDATA%\Claude\claude_desktop_config.json` in Notepad and check
  it's valid JSON (a missing comma from manual editing is the usual culprit).
  `setup.mjs` writes valid JSON, so this is more likely if you hand-edited it
  afterward.
- Confirm `dist\index.js` actually exists — `npm run build` must have
  succeeded (step 4).

**A tool call fails with something about a locked or busy file.**
Close `Job_Tracking.xlsx` in Excel first — the server can't write to a file
Excel currently has open.

**`npm install` fails partway through, or hangs.**
Usually network/firewall related — see the note in step 4 about
`cdn.sheetjs.com`. Retry on a different network if you suspect this.

**You want to start over cleanly.**
Delete the `job-tracker` entry from `claude_desktop_config.json` (or restore
the `.bak` file `setup.mjs` created next to it), then re-run `node setup.mjs`.

## Updating later

If you cloned with Git:

```powershell
cd C:\mcp-job-tracker
git pull
npm install
npm run build
```

Then fully restart Claude Desktop. Your data files (in Documents) are
untouched by this — only the server code changes.

## Where to go from here

See the main [README.md](README.md) for the full tool reference, the
day-to-day workflow (daily job-search routine, resume generation, interview
prep), and how the spreadsheet schemas work.
