import { spawn } from "node:child_process";
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { OfficeError, requireThat } from "../infrastructure/errors.mjs";
import { decodePackage } from "../formats/opc/ooxml.mjs";
import { packageEvidence } from "../formats/opc/engine-evidence.mjs";
import { preserveWorkbookEmbeddings } from "../formats/opc/preserve-embeddings.mjs";

export async function aspose(task, bytes) {
  if (task.kind === "office_engine_status" && !task.engines?.python)
    return {
      aspose: { configured: false },
      superdoc: { installed: true, version: "2.2.1" },
      pdfLib: { installed: true, version: "1.17.1" },
    };
  requireThat(
    task.engines?.python,
    "ENGINE_UNAVAILABLE",
    "Configure engines.python to an interpreter with aspose-cells-python 26.9.0 and aspose-slides 26.8.1 installed.",
  );
  const inputPackage = bytes ? decodePackage(bytes, task.maxExpandedBytes) : null;
  const resultFile = path.join(task.scratch, "result.json");
  await new Promise((resolve, reject) => {
    const child = spawn(
      task.engines.python,
      [fileURLToPath(new URL("./aspose-worker.py", import.meta.url))],
      {
        stdio: ["pipe", "ignore", "ignore"],
        env: {
          ...process.env,
          PYTHONDONTWRITEBYTECODE: "1",
          XDG_CACHE_HOME: task.scratch,
        },
      },
    );
    child.once("error", () =>
      reject(
        new OfficeError(
          "ENGINE_UNAVAILABLE",
          "Configured Aspose interpreter could not start.",
        ),
      ),
    );
    child.once("exit", (code) =>
      code === 0
        ? resolve()
        : reject(
            new OfficeError(
              "ENGINE_FAILED",
              "Aspose runtime exited without a result.",
            ),
          ),
    );
    child.stdin.on("error", () => {});
    child.stdin.end(JSON.stringify({ ...task, resultFile }));
  });
  const response = JSON.parse(readFileSync(resultFile, "utf8"));
  if (!response.ok)
    throw new OfficeError(response.error.code, response.error.message);
  if (task.kind === "office_engine_status")
    return {
      ...response.result,
      superdoc: { installed: true, version: "2.2.1" },
      pdfLib: { installed: true, version: "1.17.1" },
    };
  if (task.kind.endsWith("_engine_read"))
    return response.result;
  let outputPackage = decodePackage(readFileSync(task.output), task.maxExpandedBytes);
  if (task.kind.startsWith("xlsx_")) {
    const restored = preserveWorkbookEmbeddings(
      inputPackage,
      outputPackage,
      task.maxExpandedBytes,
    );
    if (restored.restoredEmbeddings) writeFileSync(task.output, restored.bytes);
    outputPackage = restored.packageContext;
    response.result.restoredEmbeddings = restored.restoredEmbeddings;
  }
  const evidence = packageEvidence(
    inputPackage,
    outputPackage,
    task.maxExpandedBytes,
  );
  return {
    ...response.result,
    revision: evidence.outputRevision,
    engineEvidence: { ...evidence, ...response.result },
  };
}
