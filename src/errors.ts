/**
 * Shared error types, fs-error mapping, and MCP result helpers.
 */

import * as path from "node:path";

/* ------------------------------------------------------------------ */
/* Errors                                                             */
/* ------------------------------------------------------------------ */

/** An error whose message is safe and useful to show the user directly. */
export class UserFacingError extends Error {}

/* ------------------------------------------------------------------ */
/* fs error mapping                                                   */
/* ------------------------------------------------------------------ */

export function wrapFsError(
  err: unknown,
  action: string,
  filePath: string,
  /** What the file is, for the message. Defaults to the spreadsheet wording. */
  noun: "spreadsheet" | "document" = "spreadsheet"
): UserFacingError {
  const e = err as NodeJS.ErrnoException;
  const isSheet = noun === "spreadsheet";
  if (e && (e.code === "EBUSY" || e.code === "EPERM" || e.code === "EACCES")) {
    return new UserFacingError(
      `Could not ${action} the ${noun} — it looks like ${path.basename(
        filePath
      )} is open in ${isSheet ? "Excel" : "another program"} (or locked by ` +
        `another program). Please close it and try again.`
    );
  }
  if (e && e.code === "ENOENT") {
    const label = isSheet ? "Spreadsheet" : "File";
    return new UserFacingError(`${label} not found at:\n  ${filePath}`);
  }
  return new UserFacingError(
    `Could not ${action} the ${noun}: ${e?.message ?? String(err)}`
  );
}

/* ------------------------------------------------------------------ */
/* Presentation                                                       */
/* ------------------------------------------------------------------ */

export function textResult(payload: unknown) {
  const text =
    typeof payload === "string" ? payload : JSON.stringify(payload, null, 2);
  return { content: [{ type: "text" as const, text }] };
}

export function errorResult(message: string) {
  return {
    content: [{ type: "text" as const, text: message }],
    isError: true as const,
  };
}

/** Run a tool body, converting thrown errors into clean MCP error results. */
export async function guard<T>(fn: () => T | Promise<T>) {
  try {
    return await fn();
  } catch (err) {
    if (err instanceof UserFacingError) return errorResult(err.message);
    return errorResult(
      `Unexpected error: ${err instanceof Error ? err.message : String(err)}`
    );
  }
}
