// Explicit maintenance tool. Run against an archived prior release, never current code.
// node scripts/capture-state-fixture.mjs /absolute/prior/OfficeMCP /absolute/output.json <git-commit>
import fs from "node:fs";
import path from "node:path";
import os from "node:os";
import { pathToFileURL } from "node:url";
const [prior, output, commit] = process.argv.slice(2);
if (!prior || !output || !commit) throw new Error("Prior source directory, output fixture and source commit are required.");
const load = file => import(pathToFileURL(path.join(prior, file)).href);
const { OfficeService } = await load("src/service.mjs");
const { configFrom, digest } = await load("src/files.mjs");
const { writeFixtures } = await load("test/fixtures.mjs");
const root = fs.mkdtempSync(path.join(os.tmpdir(), "office-historical-state-"));
const docs = path.join(root, "docs"); fs.mkdirSync(docs); writeFixtures(docs);
const config = path.join(root, "config.json");
fs.writeFileSync(config, JSON.stringify({ readRoots: [docs], writeRoots: { output: docs }, stateDirectory: path.join(root, "state") }));
const service = new OfficeService(configFrom(config));
const calls = [];
async function call(name, args) {
  const response = await service.call(name, args);
  if (response.isError) throw new Error(JSON.stringify(response.payload));
  calls.push({ name, args, payload: response.payload });
  return response.payload;
}
try {
  const { document } = await call("office_open", { path: path.join(docs, "contract.docx") });
  const edit = { documentId: document.id, expectedSourceRevision: document.sourceRevision,
    operationId: "historical-edit", paragraphId: "p_0", expectedText: "30 days", replacement: "45 days",
    expectedMatches: 1, changeMode: "direct", preservationProfile: "docx.text-replacement.v1" };
  const { data: candidate } = await call("docx_replace_text", edit);
  const candidateArgs = { documentId: document.id, candidateId: candidate.candidateId, expectedCandidateRevision: candidate.candidateRevision };
  const { data: validation } = await call("office_validate", candidateArgs);
  await call("office_commit", { ...candidateArgs, expectedSourceRevision: document.sourceRevision,
    operationId: "historical-save", validationId: validation.validationId, mode: "save_as", overwrite: false,
    destination: { rootId: "output", relativePath: "published.docx" } });
  const records = service.db.prepare("SELECT kind,id,value FROM records ORDER BY kind,id").all()
    .filter(row => row.kind !== "owner").map(row => ({ kind: row.kind, id: row.id, value: JSON.parse(row.value) }));
  const paths = new Set(records.flatMap(row => [row.value.source, row.value.snapshot, row.value.path]).filter(Boolean));
  const files = [...paths].filter(file => fs.existsSync(file)).map(file => {
    const bytes = fs.readFileSync(file);
    return { path: path.relative(root, file), sha256: digest(bytes), base64: bytes.toString("base64") };
  });
  const fixture = { provenance: { version: JSON.parse(fs.readFileSync(path.join(prior, "package.json"))).version,
    commit, node: process.versions.node, capturedAt: new Date().toISOString(),
    method: "Produced by executing the archived prior-release OfficeService with its pinned SuperDoc dependency; paths normalized only for relocation." },
    records, calls, files };
  fs.writeFileSync(output, JSON.stringify(fixture, null, 2).split(root).join("$ROOT") + "\n");
  console.log(JSON.stringify({ version: fixture.provenance.version, records: records.length, files: files.length, output }));
} finally { service.close(); fs.rmSync(root, { recursive: true, force: true }); }
