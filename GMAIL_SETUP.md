# Gmail setup

The Gmail tools (`search_gmail_for_job`, `read_gmail_message`, `scan_job_updates`,
`draft_gmail_reply`, `send_gmail_email`) need a Google OAuth client and a one-time
consent flow.
The first part — creating the OAuth client — has to be done by hand in the
Google Cloud Console; it can't be scripted. The second part is a single
command.

## 1. Create a Google Cloud project

1. Go to https://console.cloud.google.com/projectcreate and create a new
   project (any name, e.g. "Job Tracker Gmail").
2. Make sure the new project is selected in the project picker at the top of
   the console.

## 2. Enable the Gmail API

1. Go to https://console.cloud.google.com/apis/library/gmail.googleapis.com
2. Click **Enable** (with the new project selected).

## 3. Configure the OAuth consent screen

1. Go to https://console.cloud.google.com/apis/credentials/consent
2. Choose **External** (unless you have a Google Workspace org and want
   **Internal**), then fill in the required fields (app name, your email as
   support/developer contact). You don't need to submit for verification —
   an unverified app works fine in "Testing" mode for your own account.
3. On the Scopes step, you can skip adding scopes here (the auth script
   requests them directly); just continue through.
4. On the Test users step, add your own Gmail address so you're allowed to
   sign in while the app is unverified.
5. Save.

## 4. Create an OAuth client

1. Go to https://console.cloud.google.com/apis/credentials
2. Click **Create Credentials > OAuth client ID**.
3. Application type: **Desktop app**. Name it anything (e.g. "job-tracker CLI").
4. Click **Create**, then **Download JSON** on the client you just created.
5. Save that downloaded file as `credentials.json` directly in this project's
   root folder — the same folder as `package.json`
   (`C:\mcp-job-tracker\credentials.json`).

`credentials.json` and the token file it produces are already listed in
`.gitignore`, so they won't get committed.

## 5. Run the one-time consent flow

From this project's folder:

```
npm run gmail:auth
```

This opens your browser to Google's sign-in/consent screen (scopes: read
Gmail, compose drafts, send mail), then saves the resulting token to
`gmail_token.json` next to `credentials.json`. You only need to do this once;
the server refreshes the token automatically after that.

If you ever revoke access (https://myaccount.google.com/permissions) or want
to redo consent, just delete `gmail_token.json` and run `npm run gmail:auth`
again.

## 6. Restart Claude Desktop

Fully quit and reopen Claude Desktop so the rebuilt server (and its new Gmail
tools) get picked up.

## Notes on scope / risk

- `search_gmail_for_job` and `scan_job_updates` are read-only.
- `draft_gmail_reply` creates a real Gmail draft but never sends it.
- `send_gmail_email` sends immediately and can't be recalled — it requires an
  explicit `confirm: true` and is meant to only be used after you've reviewed
  the exact text.
- The requested scopes (`gmail.readonly`, `gmail.compose`, `gmail.send`)
  intentionally do **not** include full mailbox modify/delete access.
