import { readdirSync, writeFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import {fileURLToPath} from "node:url";
import { catalog } from "../src/contracts/catalog.mjs";
import { manifest } from "../src/contracts/registry.mjs";

function sourceModules(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const target = new URL(entry.name + (entry.isDirectory() ? "/" : ""), directory);
    return entry.isDirectory()
      ? sourceModules(target)
      : entry.name.endsWith(".mjs")
        ? [target]
        : [];
  });
}

for (const file of sourceModules(new URL("../src/", import.meta.url)))
  execFileSync(
    process.execPath,
    ["--check", fileURLToPath(file)],
    { stdio: "inherit" },
  );
writeFileSync(
  new URL("../contracts/tools.json", import.meta.url),
  JSON.stringify(manifest(), null, 2) + "\n",
);
writeFileSync(
  new URL("../contracts/catalog.json", import.meta.url),
  JSON.stringify(catalog(), null, 2) + "\n",
);
writeFileSync(
  new URL("../contracts/full-tools.json", import.meta.url),
  JSON.stringify(manifest("full"), null, 2) + "\n",
);
const docxAgentManifest = manifest("docx-agent");
writeFileSync(
  new URL("../contracts/docx-agent-tools.json", import.meta.url),
  JSON.stringify(docxAgentManifest, null, 2) + "\n",
);
await import("./dependency-inventory.mjs");
console.log("Built executable ESM sources and public tool schemas.");
