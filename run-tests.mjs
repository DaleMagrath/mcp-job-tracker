/**
 * Runs both test suites (integration + unit) as one command, so `npm test`
 * and `npm run coverage` see a single process to wrap — necessary for c8 to
 * collect coverage from both, since NODE_V8_COVERAGE only propagates to
 * child processes spawned by the one process c8 directly wraps.
 */
import { execFileSync } from "node:child_process";

for (const file of ["test-client.mjs", "test-unit.mjs"]) {
  execFileSync(process.execPath, [file], { stdio: "inherit" });
}
