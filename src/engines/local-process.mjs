import { spawn } from "node:child_process";
import { OfficeError } from "../infrastructure/errors.mjs";

/** Execute an operator-configured local converter or text extractor. */
export function runLocal(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: ["ignore", "pipe", "ignore"], ...options });
    const chunks = [];
    let size = 0, done = false;
    const finish = (error) => {
      if (done) return;
      done = true;
      error ? reject(error) : resolve(Buffer.concat(chunks).toString("utf8"));
    };
    child.stdout.on("data", (chunk) => {
      if (done) return;
      size += chunk.length;
      if (size > 2_000_000) {
        child.kill("SIGKILL");
        finish(new OfficeError("LIMIT_EXCEEDED", "Local engine output exceeded its budget."));
      } else chunks.push(chunk);
    });
    child.once("error", () => finish(new OfficeError("ENGINE_UNAVAILABLE", "Local engine executable could not start.")));
    child.once("close", (code) => finish(code === 0 ? null : new OfficeError(
      "ENGINE_FAILED", "Local engine failed. Check installed dependencies and language packs.",
    )));
  });
}
