#!/usr/bin/env node
/**
 * One-time interactive Gmail OAuth consent flow.
 *
 * Prerequisite (manual, can't be scripted): create a Google Cloud project,
 * enable the Gmail API, configure the OAuth consent screen, and create a
 * "Desktop app" OAuth client — see GMAIL_SETUP.md. Download that client's
 * JSON and save it as credentials.json in this project's root (next to
 * package.json).
 *
 * What this script does:
 *   1. Reads client_id/secret from credentials.json.
 *   2. Opens your default browser to Google's consent screen, requesting
 *      read + compose + send Gmail scopes.
 *   3. Runs a tiny local HTTP server on a loopback port to catch the
 *      redirect with the authorization code — nothing leaves your machine.
 *   4. Exchanges the code for tokens and saves them to gmail_token.json.
 *
 * Run again any time to re-consent (e.g. after revoking access in your
 * Google Account, or after adding a scope) — it always requests a fresh
 * refresh_token (prompt=consent).
 *
 * Usage: npm run gmail:auth
 */

import { exec } from "node:child_process";
import * as fs from "node:fs";
import * as http from "node:http";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import { OAuth2Client } from "google-auth-library";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.join(__dirname, "..");

const CREDENTIALS_FILE =
  (process.env.JOB_GMAIL_CREDENTIALS_FILE || "").trim() ||
  path.join(PROJECT_ROOT, "credentials.json");
const TOKEN_FILE =
  (process.env.JOB_GMAIL_TOKEN_FILE || "").trim() ||
  path.join(PROJECT_ROOT, "gmail_token.json");

const SCOPES = [
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.send",
];

const REDIRECT_PORT = 53682;
const REDIRECT_URI = `http://127.0.0.1:${REDIRECT_PORT}/oauth2callback`;

function fail(msg) {
  console.error("\n[gmail-authorize] ERROR: " + msg);
  process.exit(1);
}

if (!fs.existsSync(CREDENTIALS_FILE)) {
  fail(
    `credentials.json not found at:\n  ${CREDENTIALS_FILE}\n\n` +
      "Follow GMAIL_SETUP.md to create a Google Cloud OAuth client (Desktop app) " +
      "and download it there first."
  );
}

let raw;
try {
  raw = JSON.parse(fs.readFileSync(CREDENTIALS_FILE, "utf8"));
} catch (err) {
  fail(`${CREDENTIALS_FILE} is not valid JSON: ${err.message}`);
}

const section = raw.installed || raw.web || raw;
if (!section?.client_id || !section?.client_secret) {
  fail(
    `${CREDENTIALS_FILE} doesn't look like a Google OAuth client file (missing ` +
      "client_id/client_secret). Re-download it from the Google Cloud Console."
  );
}

const oAuth2Client = new OAuth2Client(
  section.client_id,
  section.client_secret,
  REDIRECT_URI
);

const authUrl = oAuth2Client.generateAuthUrl({
  access_type: "offline", // required to get a refresh_token back
  prompt: "consent", // force a fresh refresh_token even on repeat runs
  scope: SCOPES,
});

function openBrowser(url) {
  const platform = process.platform;
  const cmd =
    platform === "win32"
      ? `start "" "${url}"`
      : platform === "darwin"
      ? `open "${url}"`
      : `xdg-open "${url}"`;
  exec(cmd, (err) => {
    if (err) {
      console.log(
        "[gmail-authorize] Could not auto-open a browser — copy the URL above into one."
      );
    }
  });
}

function waitForAuthCode() {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      let reqUrl;
      try {
        reqUrl = new URL(req.url, REDIRECT_URI);
      } catch {
        res.writeHead(400);
        res.end();
        return;
      }
      if (reqUrl.pathname !== "/oauth2callback") {
        res.writeHead(404);
        res.end();
        return;
      }
      const code = reqUrl.searchParams.get("code");
      const error = reqUrl.searchParams.get("error");
      if (error) {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(
          `<html><body><h2>Authorization failed</h2><p>${error}</p>` +
            "<p>You can close this tab and check the terminal.</p></body></html>"
        );
        server.close();
        reject(new Error(`Google returned an error: ${error}`));
        return;
      }
      res.writeHead(200, { "Content-Type": "text/html" });
      res.end(
        "<html><body><h2>Gmail authorized</h2>" +
          "<p>You can close this tab and go back to the terminal.</p></body></html>"
      );
      server.close();
      resolve(code);
    });
    server.on("error", reject);
    server.listen(REDIRECT_PORT, "127.0.0.1", () => {
      console.log(
        "\n[gmail-authorize] Opening your browser for Google sign-in/consent..."
      );
      console.log("If it doesn't open automatically, visit this URL:\n");
      console.log(authUrl + "\n");
      openBrowser(authUrl);
    });
  });
}

async function main() {
  const code = await waitForAuthCode();
  const { tokens } = await oAuth2Client.getToken(code);
  if (!tokens.refresh_token) {
    fail(
      "Google didn't return a refresh_token. This usually means this Google " +
        "account already granted consent without 'prompt=consent' sticking — " +
        "go to https://myaccount.google.com/permissions, remove access for this " +
        "app, and run 'npm run gmail:auth' again."
    );
  }
  fs.writeFileSync(TOKEN_FILE, JSON.stringify(tokens, null, 2) + "\n", "utf8");
  console.log(`[gmail-authorize] Success. Saved tokens to:\n  ${TOKEN_FILE}`);
  console.log(
    "Restart Claude Desktop (or just try a Gmail tool) — Gmail access is ready.\n"
  );
}

main().catch((err) => fail(err?.message || String(err)));
