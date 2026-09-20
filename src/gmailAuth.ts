/**
 * Gmail OAuth2 client + a thin REST wrapper.
 *
 * Client credentials (client_id/secret) come from a Google Cloud OAuth
 * "Desktop app" client, downloaded as credentials.json — a one-time manual
 * step in the Google Cloud Console (see GMAIL_SETUP.md); it can't be
 * scripted. The actual user consent + refresh token is obtained once via
 * `npm run gmail:auth` (scripts/gmail-authorize.mjs), which saves
 * gmail_token.json. This module loads both and hands back an authorized
 * OAuth2Client (from google-auth-library), persisting any refreshed access
 * token back to disk so later calls don't need to re-consent.
 *
 * Deliberately uses only google-auth-library (small, just the auth client)
 * plus plain REST calls against the Gmail API — not the full `googleapis`
 * package, which bundles typed clients for hundreds of unrelated Google
 * APIs and is far heavier than this server needs.
 */

import * as fs from "node:fs";
import { OAuth2Client } from "google-auth-library";
import { UserFacingError } from "./errors.js";
import {
  GMAIL_CREDENTIALS_FILE,
  GMAIL_TOKEN_FILE,
  GMAIL_REDIRECT_URI,
} from "./gmailConfig.js";

interface StoredClientCredentials {
  client_id: string;
  client_secret: string;
}

function loadClientCredentials(): StoredClientCredentials {
  if (!fs.existsSync(GMAIL_CREDENTIALS_FILE)) {
    throw new UserFacingError(
      `No Gmail OAuth client found at:\n  ${GMAIL_CREDENTIALS_FILE}\n` +
        `Follow GMAIL_SETUP.md to create a Google Cloud OAuth client and download ` +
        `credentials.json there, then run "npm run gmail:auth" from the project folder.`
    );
  }
  let raw: any;
  try {
    raw = JSON.parse(fs.readFileSync(GMAIL_CREDENTIALS_FILE, "utf8"));
  } catch (err: any) {
    throw new UserFacingError(
      `${GMAIL_CREDENTIALS_FILE} is not valid JSON: ${err?.message || String(err)}`
    );
  }
  // Google's downloaded client JSON nests fields under "installed" (Desktop
  // app) or "web" (Web app); tolerate either, or a flat shape.
  const section = raw.installed || raw.web || raw;
  if (!section?.client_id || !section?.client_secret) {
    throw new UserFacingError(
      `${GMAIL_CREDENTIALS_FILE} doesn't look like a Google OAuth client file ` +
        `(missing client_id/client_secret). Re-download it from the Google Cloud ` +
        `Console under APIs & Services > Credentials.`
    );
  }
  return { client_id: section.client_id, client_secret: section.client_secret };
}

function newOAuth2Client(): OAuth2Client {
  const { client_id, client_secret } = loadClientCredentials();
  return new OAuth2Client(client_id, client_secret, GMAIL_REDIRECT_URI);
}

/**
 * Returns an OAuth2 client with the stored token loaded. The client
 * refreshes the access token automatically when it's expired (using the
 * stored refresh_token); the "tokens" listener below persists whatever
 * comes back so the next call doesn't need to refresh again.
 */
export function getAuthorizedOAuth2Client(): OAuth2Client {
  if (!fs.existsSync(GMAIL_TOKEN_FILE)) {
    throw new UserFacingError(
      `Gmail isn't authorized yet — no token at:\n  ${GMAIL_TOKEN_FILE}\n` +
        `Run "npm run gmail:auth" once from the project folder to grant access, ` +
        `then try again.`
    );
  }
  let tokens: Record<string, unknown>;
  try {
    tokens = JSON.parse(fs.readFileSync(GMAIL_TOKEN_FILE, "utf8"));
  } catch (err: any) {
    throw new UserFacingError(
      `${GMAIL_TOKEN_FILE} is not valid JSON. Delete it and re-run "npm run gmail:auth".`
    );
  }
  if (!tokens.refresh_token) {
    throw new UserFacingError(
      `${GMAIL_TOKEN_FILE} has no refresh_token — the consent flow needs to be ` +
        `redone. Delete the file and re-run "npm run gmail:auth".`
    );
  }

  const client = newOAuth2Client();
  client.setCredentials(tokens);

  client.on("tokens", (newTokens) => {
    try {
      const merged = { ...tokens, ...newTokens };
      fs.writeFileSync(GMAIL_TOKEN_FILE, JSON.stringify(merged, null, 2) + "\n", "utf8");
    } catch {
      /* best-effort persistence; the in-memory client still works this call */
    }
  });

  return client;
}

const GMAIL_API_BASE = "https://gmail.googleapis.com/gmail/v1/users/me";

/** Build a URL with query params, repeating array values (Gmail's convention
 *  for e.g. metadataHeaders=From&metadataHeaders=Subject). */
export function buildUrl(path: string, params: Record<string, unknown> | undefined): string {
  const qs = new URLSearchParams();
  for (const [k, v] of Object.entries(params ?? {})) {
    if (v === undefined || v === null) continue;
    if (Array.isArray(v)) {
      for (const item of v) qs.append(k, String(item));
    } else {
      qs.append(k, String(v));
    }
  }
  const query = qs.toString();
  return `${GMAIL_API_BASE}${path}${query ? `?${query}` : ""}`;
}

/** Wrap Gmail REST errors (which arrive as thrown GaxiosError) into a UserFacingError. */
export function wrapGmailError(err: any, action: string): never {
  const status = err?.response?.status ?? err?.code;
  const apiMessage = err?.response?.data?.error?.message;
  if (status === 401 || status === 403) {
    throw new UserFacingError(
      `Gmail refused the request (${status}) while trying to ${action}${
        apiMessage ? `: ${apiMessage}` : ""
      }. The token may be expired or revoked — re-run "npm run gmail:auth".`
    );
  }
  throw new UserFacingError(
    `Gmail API error while trying to ${action}${apiMessage ? `: ${apiMessage}` : `: ${err?.message || String(err)}`}`
  );
}

/* c8 ignore start -- these three functions each make a real HTTPS call to
 * the Gmail API via an authorized client; the test suite never runs with
 * real Gmail credentials (see gmailTools.ts's "no real inbox" test notes),
 * so they're structurally unreachable there. Pure helpers above (buildUrl,
 * wrapGmailError, loadClientCredentials's parsing) are NOT excluded. */
/** GET against the Gmail API (users/me/...), returning the parsed JSON body. */
export async function gmailGet<T = any>(
  auth: OAuth2Client,
  path: string,
  params?: Record<string, unknown>
): Promise<T> {
  try {
    const res = await auth.request<T>({ url: buildUrl(path, params) });
    return res.data;
  } catch (err) {
    wrapGmailError(err, `read ${path}`);
  }
}

/** POST against the Gmail API (users/me/...), returning the parsed JSON body. */
export async function gmailPost<T = any>(
  auth: OAuth2Client,
  path: string,
  body: unknown
): Promise<T> {
  try {
    const res = await auth.request<T>({
      url: `${GMAIL_API_BASE}${path}`,
      method: "POST",
      data: body,
    });
    return res.data;
  } catch (err) {
    wrapGmailError(err, `write to ${path}`);
  }
}

/** The authenticated account's own email address. */
export async function getGmailAddress(): Promise<string> {
  const auth = getAuthorizedOAuth2Client();
  const data = await gmailGet<{ emailAddress?: string }>(auth, "/profile");
  return data.emailAddress ?? "";
}
/* c8 ignore stop */
