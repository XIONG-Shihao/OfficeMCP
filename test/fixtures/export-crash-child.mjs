import fs from "node:fs";
import { OfficeService } from "../../src/application/service.mjs";
import { configFrom } from "../../src/infrastructure/files.mjs";
const [config, phase, request] = process.argv.slice(2);
const service = new OfficeService(configFrom(config), {
  checkpoint(name) { if (name === phase) process.kill(process.pid, "SIGKILL"); },
});
const { args, file } = JSON.parse(fs.readFileSync(request, "utf8"));
await service.documents.exportArtifact("docx_export", args, "pdf", () => ({ file, mediaType: "application/pdf", engine: "fixture" }));
service.close();
process.exitCode = 9;
