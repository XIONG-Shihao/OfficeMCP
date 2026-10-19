import { readFileSync } from "node:fs";
import { pathToFileURL } from "node:url";
import path from "node:path";
const base = process.env.RPF_TEST_OFFICEMCP_ROOT
  ? pathToFileURL(path.resolve(process.env.RPF_TEST_OFFICEMCP_ROOT) + path.sep)
  : new URL("../", import.meta.url);
const [{ OfficeService }, { configFrom }] = await Promise.all([
  import(new URL("src/application/service.mjs", base).href),
  import(new URL("src/infrastructure/files.mjs", base).href),
]);
const [cfg, phase, request] = process.argv.slice(2);
const config = configFrom(cfg);
const publication = process.env.RPF_TEST_PUBLICATION_ADAPTER ? await (await import(pathToFileURL(process.env.RPF_TEST_PUBLICATION_ADAPTER).href)).createPublication(config) : null;
const s = new OfficeService(config, {
  checkpoint(name) {
    if (name === phase) process.kill(process.pid, "SIGKILL");
  },
}, { publication });
if (request) {
  const r = JSON.parse(readFileSync(request, "utf8"));
  await s.call(r.name, r.args);
}
s.close();
process.exitCode = 9; // Reaching here means the requested crash boundary was not exercised.
