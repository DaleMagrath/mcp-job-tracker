#!/usr/bin/env node
/**
 * One-shot installer for the job-tracker MCP server on a new machine.
 *
 * What it does:
 *   1. Runs `npm install` and builds the TypeScript (unless --no-build).
 *   2. Locates every plausible Claude Desktop config on this machine (Windows /
 *      macOS / Linux) — on Windows, that includes a Store-installed (MSIX)
 *      Claude Desktop's sandboxed config location, not just the traditional
 *      %APPDATA% one (see candidateConfigPaths() below for why both matter).
 *   3. Backs each one up, then merges in a "job-tracker" mcpServers entry with
 *      absolute paths that are correct for THIS machine.
 *
 * Usage:
 *   node setup.mjs [path-to-Job_Tracking.xlsx] [options]
 *
 * Options:
 *   --no-build      Skip npm install / build (use the dist/ you copied over).
 *   --bare-node     Use "node" in the config instead of this exact node binary.
 *   -h, --help      Show this help.
 *
 * If you don't pass a spreadsheet path, it defaults to
 *   <home>/Documents/Job_Tracking.xlsx
 */

import { spawnSync } from "node:child_process";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const argv = process.argv.slice(2);

if (argv.includes("-h") || argv.includes("--help")) {
  console.log(
    [
      "Usage: node setup.mjs [path-to-Job_Tracking.xlsx] [--no-build] [--bare-node]",
      "",
      "Installs dependencies, builds, and wires this project into Claude Desktop's",
      "config on this machine. Backs up the existing config first.",
    ].join("\n")
  );
  process.exit(0);
}

const noBuild = argv.includes("--no-build");
const bareNode = argv.includes("--bare-node");
const positional = argv.filter((a) => !a.startsWith("--"));

function fail(msg) {
  console.error("\n[setup] ERROR: " + msg);
  process.exit(1);
}

/* 1. Install + build ------------------------------------------------ */
if (!noBuild) {
  // Pass the whole command as one shell string (no args array) so Node doesn't
  // emit the DEP0190 "args + shell:true" deprecation warning. Args are static.
  console.log("[setup] Installing dependencies (npm install)...");
  let r = spawnSync("npm install", { cwd: __dirname, stdio: "inherit", shell: true });
  if (r.status !== 0) fail("`npm install` failed. Is Node/npm installed and online?");

  console.log("[setup] Building (npm run build)...");
  r = spawnSync("npm run build", { cwd: __dirname, stdio: "inherit", shell: true });
  if (r.status !== 0) fail("`npm run build` failed. See the tsc output above.");
} else {
  console.log("[setup] --no-build given; skipping install/build.");
}

const distEntry = path.join(__dirname, "dist", "index.js");
if (!fs.existsSync(distEntry)) {
  fail(
    `Built server not found at ${distEntry}. ` +
      (noBuild ? "Run without --no-build to build it." : "The build did not produce it.")
  );
}

/* 2. Resolve the spreadsheet + config paths ------------------------- */
const xlsxPath = path.resolve(
  positional[0] || path.join(os.homedir(), "Documents", "Job_Tracking.xlsx")
);
if (!fs.existsSync(xlsxPath)) {
  console.warn(
    `[setup] NOTE: spreadsheet not found at ${xlsxPath}.\n` +
      "       The server will report a clear error on most tools until it exists there —\n" +
      "       either copy an existing Job_Tracking.xlsx into place, or call this server's\n" +
      "       init_job_tracker_files tool once to create a blank one at that path.\n" +
      "       Re-run with the correct path if this is wrong: node setup.mjs <path>"
  );
}

/**
 * Every plausible Claude Desktop config location on this machine.
 *
 * On Windows, a Store-installed (MSIX/packaged) Claude Desktop runs inside an
 * app container: when it reads/writes what it thinks is
 * "%APPDATA%\Claude\claude_desktop_config.json", Windows silently redirects
 * that to an isolated per-package folder instead
 * ("%LOCALAPPDATA%\Packages\Claude_<hash>\LocalCache\Roaming\Claude\..."). A
 * plain script like this one has no such redirection, so writing only the
 * traditional path leaves a Store install with no entry at all — the config
 * gets written, just not to the file the app actually reads. Detect any such
 * package folder (its existence alone proves a Store install is present,
 * even before it has ever created its own config file) and write to it too.
 */
function candidateConfigPaths() {
  if (process.platform === "win32") {
    const appData = process.env.APPDATA || path.join(os.homedir(), "AppData", "Roaming");
    const paths = [path.join(appData, "Claude", "claude_desktop_config.json")];

    const localAppData = process.env.LOCALAPPDATA || path.join(os.homedir(), "AppData", "Local");
    const packagesDir = path.join(localAppData, "Packages");
    try {
      for (const name of fs.readdirSync(packagesDir)) {
        if (/^Claude_/i.test(name)) {
          paths.push(
            path.join(packagesDir, name, "LocalCache", "Roaming", "Claude", "claude_desktop_config.json")
          );
        }
      }
    } catch {
      /* no Packages dir, or unreadable — not a Store-app-aware machine, or none installed */
    }
    return paths;
  }
  if (process.platform === "darwin") {
    return [
      path.join(os.homedir(), "Library", "Application Support", "Claude", "claude_desktop_config.json"),
    ];
  }
  // Linux / other (Claude Desktop isn't officially here, but be predictable).
  return [path.join(os.homedir(), ".config", "Claude", "claude_desktop_config.json")];
}

const cfgPaths = candidateConfigPaths();
if (cfgPaths.length > 1) {
  console.log(`[setup] Found ${cfgPaths.length} possible Claude Desktop config locations on this machine —`);
  console.log("        writing the job-tracker entry into each, since only one is actually read:");
  for (const p of cfgPaths) console.log(`          ${p}`);
}

/* 3. Merge into the config(s) (backing up first) --------------------- */
function mergeInto(cfgPath) {
  let cfg = {};
  if (fs.existsSync(cfgPath)) {
    const raw = fs.readFileSync(cfgPath, "utf8");
    try {
      cfg = raw.trim() ? JSON.parse(raw) : {};
    } catch {
      fail(
        `${cfgPath} is not valid JSON. Refusing to overwrite it — please fix or ` +
          "remove it, then re-run."
      );
    }
    fs.copyFileSync(cfgPath, cfgPath + ".bak");
    console.log(`[setup] Backed up existing config -> ${cfgPath}.bak`);
  } else {
    fs.mkdirSync(path.dirname(cfgPath), { recursive: true });
    console.log(`[setup] No existing config; creating ${cfgPath}`);
  }

  cfg.mcpServers = cfg.mcpServers || {};
  cfg.mcpServers["job-tracker"] = {
    command: bareNode ? "node" : process.execPath,
    args: [distEntry],
    env: { JOB_TRACKER_FILE: xlsxPath },
  };

  fs.writeFileSync(cfgPath, JSON.stringify(cfg, null, 2) + "\n", "utf8");
}

for (const cfgPath of cfgPaths) mergeInto(cfgPath);

console.log("\n[setup] Done. Wired 'job-tracker' into Claude Desktop:");
console.log("        server : " + distEntry);
console.log("        node   : " + (bareNode ? "node (from PATH)" : process.execPath));
console.log("        data   : " + xlsxPath);
for (const cfgPath of cfgPaths) console.log("        config : " + cfgPath);
console.log("\n>> Fully quit and restart Claude Desktop to load the tools. <<");
