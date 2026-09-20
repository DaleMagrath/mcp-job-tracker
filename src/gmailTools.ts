/**
 * Gmail tools: search, cross-reference with the tracker, draft, and send.
 *
 * Auth: see gmailAuth.ts / GMAIL_SETUP.md. Every tool here throws a clear
 * UserFacingError (via the shared `guard` helper) pointing at "npm run
 * gmail:auth" if Gmail isn't authorized yet, so a first-time user gets a
 * useful message instead of a raw API error.
 *
 * Safety posture, deliberately graduated:
 *   - search_gmail_for_job / scan_job_updates / read_gmail_message: read-only,
 *     no confirmation needed.
 *   - draft_gmail_reply: writes a Gmail draft, but drafts are inert — nothing
 *     is sent until a human opens Gmail and hits Send.
 *   - send_gmail_email: actually sends mail immediately and cannot be undone.
 *     Requires confirm: true; the tool refuses without it. Callers should
 *     show the user the exact to/subject/body first and only pass confirm
 *     after they've said to go ahead.
 */

import { z } from "zod";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { OAuth2Client } from "google-auth-library";
import { UserFacingError, textResult, guard } from "./errors.js";
import { getAuthorizedOAuth2Client, gmailGet, gmailPost } from "./gmailAuth.js";
import {
  openWorkbook,
  readAllRecords,
  includesCI,
} from "./workbook.js";
import { TRACKER } from "./config.js";

/* ------------------------------------------------------------------ */
/* Small Gmail helpers                                                */
/* ------------------------------------------------------------------ */

const CLOSED_STATUSES = new Set([
  "declined",
  "withdrawn",
  "closed - no longer available",
]);

export function messageLink(id: string): string {
  return `https://mail.google.com/mail/u/0/#all/${id}`;
}

/** Pull one header's value (case-insensitive) from a Gmail message payload. */
export function header(payload: any, name: string): string | undefined {
  const list: any[] = payload?.headers ?? [];
  return list.find((h) => h.name?.toLowerCase() === name.toLowerCase())?.value;
}

/* c8 ignore start -- both functions below make a real Gmail API call via
 * gmailGet(); structurally unreachable without real credentials (see
 * gmailAuth.ts's matching note). The pure header()/messageLink() helpers
 * above are NOT excluded. */
/** Fetch a message with just the metadata we need (fast, no body download). */
async function getMessageMeta(auth: OAuth2Client, id: string) {
  const data = await gmailGet<any>(auth, `/messages/${id}`, {
    format: "metadata",
    metadataHeaders: ["From", "Subject", "Date", "Message-ID", "References"],
  });
  const payload = data.payload;
  return {
    id: data.id as string,
    threadId: data.threadId as string,
    from: header(payload, "From") ?? "",
    subject: header(payload, "Subject") ?? "",
    date: header(payload, "Date") ?? "",
    messageIdHeader: header(payload, "Message-ID"),
    references: header(payload, "References"),
    snippet: (data.snippet as string) ?? "",
  };
}

/** Run a Gmail search query, returning lightweight metadata per match. */
async function searchMessages(auth: OAuth2Client, q: string, maxResults: number) {
  const data = await gmailGet<any>(auth, "/messages", { q, maxResults });
  const ids = (data.messages ?? []).map((m: any) => m.id as string);
  const out = [];
  for (const id of ids) {
    out.push(await getMessageMeta(auth, id));
  }
  return out;
}
/* c8 ignore stop */

/** Decode a Gmail API base64url body payload into a UTF-8 string. */
export function decodeBase64Url(data: string): string {
  const b64 = data.replace(/-/g, "+").replace(/_/g, "/");
  return Buffer.from(b64, "base64").toString("utf8");
}

/** Strip an HTML email body down to readable plain text (best-effort). */
export function stripHtml(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6])>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;/gi, "'")
    .replace(/&quot;/gi, '"')
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** Walk a MIME payload tree, returning the best text/plain and/or text/html part found. */
export function extractBody(payload: any): { text?: string; html?: string } {
  const result: { text?: string; html?: string } = {};
  function walk(part: any): void {
    if (!part) return;
    const mime: string = part.mimeType || "";
    if (mime === "text/plain" && part.body?.data && !result.text) {
      result.text = decodeBase64Url(part.body.data);
    } else if (mime === "text/html" && part.body?.data && !result.html) {
      result.html = decodeBase64Url(part.body.data);
    }
    if (Array.isArray(part.parts)) {
      for (const p of part.parts) walk(p);
    }
  }
  walk(payload);
  // Single-part messages have no `parts` array — the body sits on the payload itself.
  if (!result.text && !result.html && payload?.body?.data) {
    if ((payload.mimeType || "").includes("html")) {
      result.html = decodeBase64Url(payload.body.data);
    } else {
      result.text = decodeBase64Url(payload.body.data);
    }
  }
  return result;
}

/** Base64url-encode a raw RFC 2822 message the way the Gmail API expects. */
export function toBase64Url(raw: string): string {
  return Buffer.from(raw, "utf8")
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

/**
 * RFC 2047-encode a header value that contains non-ASCII characters.
 *
 * Header fields must be ASCII (RFC 5322). Writing raw UTF-8 bytes into a header
 * leaves every downstream relay and client free to guess the charset, and they
 * guess Latin-1: a subject containing an em dash went out as
 * "Job Search Discovery Ã¢Â€Â” New Leads". The body never had this problem
 * because it declares charset="UTF-8".
 *
 * Encoded-words are capped at 75 characters, so long values are split into
 * several words folded onto continuation lines. Chunking happens on character
 * boundaries — never mid-codepoint — because a split multi-byte sequence would
 * decode to a replacement character.
 */
export function encodeHeaderValue(value: string): string {
  if (!/[^\x20-\x7e]/.test(value)) return value;

  const MAX_BYTES = 45; // 45 bytes -> 60 base64 chars + "=?UTF-8?B??=" = 72
  const chunks: string[] = [];
  let current = "";
  for (const char of value) {
    const candidate = current + char;
    if (Buffer.byteLength(candidate, "utf8") > MAX_BYTES) {
      chunks.push(current);
      current = char;
    } else {
      current = candidate;
    }
  }
  if (current) chunks.push(current);

  return chunks
    .map((c) => `=?UTF-8?B?${Buffer.from(c, "utf8").toString("base64")}?=`)
    .join("\r\n ");
}

export function buildRawEmail(opts: {
  to: string;
  cc?: string;
  subject: string;
  body: string;
  inReplyTo?: string;
  references?: string;
}): string {
  const lines = [
    `To: ${opts.to}`,
    opts.cc ? `Cc: ${opts.cc}` : undefined,
    `Subject: ${encodeHeaderValue(opts.subject)}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/plain; charset="UTF-8"`,
    opts.inReplyTo ? `In-Reply-To: ${opts.inReplyTo}` : undefined,
    opts.references ? `References: ${opts.references}` : undefined,
  ].filter((l): l is string => l !== undefined);
  return toBase64Url(lines.join("\r\n") + "\r\n\r\n" + opts.body);
}

/** The most recent message in a thread — what a reply should thread onto. */
/* c8 ignore start -- makes a real Gmail API call via gmailGet() (see the
 * getMessageMeta/searchMessages note above). */
async function getLastMessageInThread(auth: OAuth2Client, threadId: string) {
  const data = await gmailGet<any>(auth, `/threads/${threadId}`, {
    format: "metadata",
    metadataHeaders: ["From", "Subject", "Message-ID", "References"],
  });
  const messages = data.messages ?? [];
  if (!messages.length) {
    throw new UserFacingError(`Thread "${threadId}" has no messages.`);
  }
  const last = messages[messages.length - 1];
  const payload = last.payload;
  return {
    from: header(payload, "From") ?? "",
    subject: header(payload, "Subject") ?? "",
    messageIdHeader: header(payload, "Message-ID"),
    references: header(payload, "References"),
  };
}
/* c8 ignore stop */

/** Extract a bare email address out of a "Name <addr@x.com>" From header. */
export function extractEmailAddress(fromHeader: string): string {
  const m = /<([^>]+)>/.exec(fromHeader);
  return (m ? m[1] : fromHeader).trim();
}

/** Simple keyword heuristics over subject+snippet, most-significant first. */
export function guessStatusFromText(text: string): string | undefined {
  const t = text.toLowerCase();
  const offerHints = ["pleased to offer", "extend an offer", "job offer", "offer letter"];
  const declineHints = [
    "unfortunately",
    "not moving forward",
    "will not be moving forward",
    "decided to move forward with other",
    "regret to inform",
    "other candidates",
    "position has been filled",
  ];
  const interviewHints = [
    "schedule a call",
    "schedule an interview",
    "phone screen",
    "next steps",
    "book a time",
    "interview",
    "chat about your background",
  ];
  if (offerHints.some((h) => t.includes(h))) return "Offer";
  if (declineHints.some((h) => t.includes(h))) return "Declined";
  if (interviewHints.some((h) => t.includes(h))) return "Interviewing";
  return undefined;
}

/* ------------------------------------------------------------------ */

export function register(server: McpServer): void {
  // G1. search_gmail_for_job -------------------------------------------
  server.registerTool(
    "search_gmail_for_job",
    {
      title: "Search Gmail for job-related email",
      description:
        "Read-only search of the authorized Gmail account for messages related " +
        "to a company (and optionally position/keywords) — recruiter replies, " +
        "interview invites, rejections, etc. Does not modify anything. Requires " +
        "Gmail to be authorized first (see GMAIL_SETUP.md / npm run gmail:auth).",
      inputSchema: {
        company: z.string().min(1).describe("Company name to search for."),
        position: z
          .string()
          .optional()
          .describe("Position/title to narrow the search further."),
        query: z
          .string()
          .optional()
          .describe(
            "Advanced override for the raw Gmail search query (Gmail search " +
              "syntax, e.g. 'from:recruiter@acme.com newer_than:30d'). If given, " +
              "replaces the default company/position query entirely."
          ),
        max_results: z
          .number()
          .int()
          .positive()
          .max(50)
          .optional()
          .describe("Max messages to return (default 10, max 50)."),
      },
    },
    async (args) =>
      guard(async () => {
        const auth = getAuthorizedOAuth2Client();
        const q =
          args.query?.trim() ||
          [args.company.trim(), args.position?.trim()].filter(Boolean).join(" ");
        const max = args.max_results ?? 10;

        const messages = await searchMessages(auth, q, max);
        return textResult({
          query: q,
          count: messages.length,
          messages: messages.map((m) => ({
            id: m.id,
            threadId: m.threadId,
            from: m.from,
            subject: m.subject,
            date: m.date,
            snippet: m.snippet,
            link: messageLink(m.id),
          })),
        });
      })
  );

  // G1b. read_gmail_message -----------------------------------------------
  server.registerTool(
    "read_gmail_message",
    {
      title: "Read a Gmail message's full body",
      description:
        "Read the full body of ONE specific Gmail message — search_gmail_for_job " +
        "and scan_job_updates only return a short snippet, so use this when you " +
        "need the actual content (e.g. an interview link or instructions). Pass " +
        "message_id (the `id` from a search/scan result), or thread_id to read " +
        "the most recent message in that thread. Read-only.",
      inputSchema: {
        message_id: z
          .string()
          .optional()
          .describe("Gmail message id, e.g. the `id` from a search_gmail_for_job result."),
        thread_id: z
          .string()
          .optional()
          .describe(
            "Gmail thread id — reads the most recent message in the thread. " +
              "Ignored if message_id is given."
          ),
      },
    },
    async (args) =>
      guard(async () => {
        const auth = getAuthorizedOAuth2Client();
        let id = args.message_id?.trim();

        if (!id) {
          const threadId = args.thread_id?.trim();
          if (!threadId) {
            throw new UserFacingError("Provide message_id or thread_id.");
          }
          const threadData = await gmailGet<any>(auth, `/threads/${threadId}`, {
            format: "metadata",
          });
          const messages = threadData.messages ?? [];
          if (!messages.length) {
            throw new UserFacingError(`Thread "${threadId}" has no messages.`);
          }
          id = messages[messages.length - 1].id;
        }

        const data = await gmailGet<any>(auth, `/messages/${id}`, { format: "full" });
        const payload = data.payload;
        const { text, html } = extractBody(payload);
        let body = text || (html ? stripHtml(html) : "") || data.snippet || "";

        const MAX = 20_000;
        let truncated = false;
        if (body.length > MAX) {
          body = body.slice(0, MAX);
          truncated = true;
        }

        return textResult({
          id: data.id,
          threadId: data.threadId,
          from: header(payload, "From") ?? "",
          to: header(payload, "To") ?? "",
          subject: header(payload, "Subject") ?? "",
          date: header(payload, "Date") ?? "",
          link: messageLink(data.id),
          ...(truncated
            ? { truncated: true, note: `Body truncated to ${MAX} characters.` }
            : {}),
          body,
        });
      })
  );

  // G2. scan_job_updates ------------------------------------------------
  server.registerTool(
    "scan_job_updates",
    {
      title: "Scan Gmail for tracker updates",
      description:
        "Cross-reference open applications in the tracker against Gmail: for " +
        "each open job (or one company, if given), searches Gmail for recent " +
        "related mail and flags simple keyword-based signals (interview, offer, " +
        "rejection language). Read-only — it only SUGGESTS a status change via " +
        "`suggestedCall`; it never updates the tracker itself. Review the emails " +
        "and call update_job_status yourself if the suggestion looks right.",
      inputSchema: {
        company: z
          .string()
          .optional()
          .describe("Limit the scan to one company (case-insensitive substring)."),
        days_back: z
          .number()
          .int()
          .positive()
          .optional()
          .describe("Only search mail from the last N days (default 30)."),
        max_per_company: z
          .number()
          .int()
          .positive()
          .max(20)
          .optional()
          .describe("Max messages to fetch per company (default 3)."),
      },
    },
    async (args) =>
      guard(async () => {
        const auth = getAuthorizedOAuth2Client();
        const h = openWorkbook(TRACKER);
        const records = readAllRecords(h);
        const daysBack = args.days_back ?? 30;
        const perCompany = args.max_per_company ?? 3;

        const open = records.filter((r) => {
          if (CLOSED_STATUSES.has((r.values["Status"] ?? "").trim().toLowerCase())) {
            return false;
          }
          return !args.company || includesCI(r.values["Company"], args.company);
        });

        if (open.length === 0) {
          return textResult({
            note: args.company
              ? `No open tracker rows match "${args.company}".`
              : "No open (non-closed) rows in the tracker to scan.",
            results: [],
          });
        }

        const results = [];
        for (const rec of open) {
          const company = rec.values["Company"];
          const position = rec.values["Position"];
          const q = `${company} newer_than:${daysBack}d`;
          const messages = await searchMessages(auth, q, perCompany);

          let suggestedStatus: string | undefined;
          for (const m of messages) {
            const guess = guessStatusFromText(`${m.subject} ${m.snippet}`);
            if (guess) {
              suggestedStatus = guess;
              break;
            }
          }

          results.push({
            company,
            position,
            currentStatus: rec.values["Status"],
            emailsFound: messages.length,
            emails: messages.map((m) => ({
              from: m.from,
              subject: m.subject,
              date: m.date,
              snippet: m.snippet,
              link: messageLink(m.id),
            })),
            ...(suggestedStatus && suggestedStatus !== rec.values["Status"]
              ? {
                  suggestedStatus,
                  suggestedCall: {
                    tool: "update_job_status",
                    arguments: { company, position, status: suggestedStatus },
                  },
                }
              : {}),
          });
        }

        return textResult({
          scanned: results.length,
          daysBack,
          results,
        });
      })
  );

  // G3. draft_gmail_reply ------------------------------------------------
  server.registerTool(
    "draft_gmail_reply",
    {
      title: "Draft a Gmail reply (not sent)",
      description:
        "Create a real Gmail draft — visible in the Drafts folder — but does " +
        "NOT send it; a human still has to open it and hit Send. Either replies " +
        "within an existing thread (pass thread_id, e.g. from search_gmail_for_job " +
        "or scan_job_updates) or starts a new message (pass to + subject). " +
        "Requires Gmail to be authorized (see GMAIL_SETUP.md).",
      inputSchema: {
        body: z.string().min(1).describe("The email body text."),
        thread_id: z
          .string()
          .optional()
          .describe(
            "Gmail thread ID to reply within (recipient/subject are inferred " +
              "from the thread's last message). Omit to start a new email."
          ),
        to: z
          .string()
          .optional()
          .describe("Recipient email address. Required if thread_id is omitted."),
        subject: z
          .string()
          .optional()
          .describe(
            "Subject line. Required if thread_id is omitted; ignored (the " +
              "thread's own subject is used) otherwise."
          ),
        cc: z.string().optional().describe("Optional Cc address(es)."),
      },
    },
    async (args) =>
      guard(async () => {
        const auth = getAuthorizedOAuth2Client();

        let to = args.to?.trim();
        let subject = args.subject?.trim();
        let inReplyTo: string | undefined;
        let references: string | undefined;
        let threadId: string | undefined = args.thread_id?.trim() || undefined;

        if (threadId) {
          const last = await getLastMessageInThread(auth, threadId);
          to = to || extractEmailAddress(last.from);
          subject = last.subject.toLowerCase().startsWith("re:")
            ? last.subject
            : `Re: ${last.subject}`;
          inReplyTo = last.messageIdHeader;
          references = [last.references, last.messageIdHeader]
            .filter(Boolean)
            .join(" ");
        } else {
          if (!to) {
            throw new UserFacingError(
              "Provide `to` (and `subject`) for a new email, or `thread_id` to reply " +
                "within an existing thread."
            );
          }
          if (!subject) {
            throw new UserFacingError(
              "Provide a `subject` for a new email (or pass thread_id to reply " +
                "within an existing thread, which reuses its subject)."
            );
          }
        }

        const raw = buildRawEmail({
          to: to!,
          cc: args.cc?.trim() || undefined,
          subject: subject!,
          body: args.body,
          inReplyTo,
          references,
        });

        const data = await gmailPost<any>(auth, "/drafts", {
          message: {
            raw,
            ...(threadId ? { threadId } : {}),
          },
        });

        return textResult({
          message: `Saved a Gmail draft to "${to}" — not sent. Open Gmail's Drafts folder to review and send it.`,
          draftId: data.id,
          to,
          subject,
          repliedToThread: threadId ?? null,
        });
      })
  );

  // G4. send_gmail_email --------------------------------------------------
  server.registerTool(
    "send_gmail_email",
    {
      title: "Send an email via Gmail",
      description:
        "SENDS a real email immediately through the authorized Gmail account. " +
        "This cannot be undone — there is no draft step. Only call this after " +
        "showing the user the exact to/subject/body and getting explicit " +
        "confirmation; the call must include confirm: true or it is refused. " +
        "Either sends a reply within an existing thread (thread_id) or a new " +
        "email (to + subject). Requires Gmail to be authorized (see GMAIL_SETUP.md).",
      inputSchema: {
        confirm: z
          .boolean()
          .describe(
            "Must be true. Set this only after the user has explicitly approved " +
              "sending this exact email."
          ),
        body: z.string().min(1).describe("The email body text."),
        thread_id: z
          .string()
          .optional()
          .describe(
            "Gmail thread ID to reply within (recipient/subject are inferred " +
              "from the thread's last message). Omit to send a new email."
          ),
        to: z
          .string()
          .optional()
          .describe("Recipient email address. Required if thread_id is omitted."),
        subject: z
          .string()
          .optional()
          .describe(
            "Subject line. Required if thread_id is omitted; ignored (the " +
              "thread's own subject is used) otherwise."
          ),
        cc: z.string().optional().describe("Optional Cc address(es)."),
      },
    },
    async (args) =>
      guard(async () => {
        if (args.confirm !== true) {
          throw new UserFacingError(
            "Refusing to send: confirm must be true. Show the user the exact " +
              "to/subject/body first and only resend this call with confirm: true " +
              "once they've said to go ahead."
          );
        }

        const auth = getAuthorizedOAuth2Client();

        let to = args.to?.trim();
        let subject = args.subject?.trim();
        let inReplyTo: string | undefined;
        let references: string | undefined;
        let threadId: string | undefined = args.thread_id?.trim() || undefined;

        if (threadId) {
          const last = await getLastMessageInThread(auth, threadId);
          to = to || extractEmailAddress(last.from);
          subject = last.subject.toLowerCase().startsWith("re:")
            ? last.subject
            : `Re: ${last.subject}`;
          inReplyTo = last.messageIdHeader;
          references = [last.references, last.messageIdHeader]
            .filter(Boolean)
            .join(" ");
        } else {
          if (!to) {
            throw new UserFacingError(
              "Provide `to` (and `subject`) for a new email, or `thread_id` to reply " +
                "within an existing thread."
            );
          }
          if (!subject) {
            throw new UserFacingError(
              "Provide a `subject` for a new email (or pass thread_id to reply " +
                "within an existing thread, which reuses its subject)."
            );
          }
        }

        const raw = buildRawEmail({
          to: to!,
          cc: args.cc?.trim() || undefined,
          subject: subject!,
          body: args.body,
          inReplyTo,
          references,
        });

        const data = await gmailPost<any>(auth, "/messages/send", {
          raw,
          ...(threadId ? { threadId } : {}),
        });

        return textResult({
          message: `Sent to "${to}".`,
          messageId: data.id,
          threadId: data.threadId,
          to,
          subject,
          link: data.id ? messageLink(data.id) : undefined,
        });
      })
  );
}
