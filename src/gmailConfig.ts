/**
 * Gmail integration configuration: credential/token file paths and scopes.
 *
 * Kept separate from config.ts (which is about the tracker/discovery/resume
 * files) because Gmail auth is independent of where the spreadsheet lives —
 * these default to living next to the server code itself (project root),
 * not next to Job_Tracking.xlsx.
 */

import * as path from "node:path";
import { SERVER_DIR } from "./config.js";

/** The project root (one level up from dist/, where package.json lives). */
export const PROJECT_ROOT = path.join(SERVER_DIR, "..");

function resolveGmailPath(envVar: string, defaultName: string): string {
  const fromEnv = process.env[envVar];
  const chosen =
    (fromEnv && fromEnv.trim()) || path.join(PROJECT_ROOT, defaultName);
  return path.resolve(chosen);
}

/**
 * OAuth client credentials downloaded from the Google Cloud Console (a
 * "Desktop app" OAuth client). See GMAIL_SETUP.md — this file can't be
 * created by a script; it's a one-time manual download.
 */
export const GMAIL_CREDENTIALS_FILE = resolveGmailPath(
  "JOB_GMAIL_CREDENTIALS_FILE",
  "credentials.json"
);

/**
 * The stored access/refresh token produced by the one-time consent flow
 * (`npm run gmail:auth`, see scripts/gmail-authorize.mjs).
 */
export const GMAIL_TOKEN_FILE = resolveGmailPath(
  "JOB_GMAIL_TOKEN_FILE",
  "gmail_token.json"
);

/**
 * Scopes requested during consent. Keep this list in sync with
 * scripts/gmail-authorize.mjs — if you add a scope here, re-run
 * `npm run gmail:auth` so the stored token actually carries it.
 */
export const GMAIL_SCOPES = [
  "https://www.googleapis.com/auth/gmail.readonly",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.send",
];

/** Loopback redirect used by the one-time consent flow. */
export const GMAIL_REDIRECT_PORT = 53682;
export const GMAIL_REDIRECT_URI = `http://127.0.0.1:${GMAIL_REDIRECT_PORT}/oauth2callback`;
