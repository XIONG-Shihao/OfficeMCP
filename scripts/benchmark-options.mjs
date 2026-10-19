import fs from "node:fs";
import path from "node:path";
import { randomUUID } from "node:crypto";

export function benchmarkOptions(project, env = process.env) {
  const requested = env.OFFICEMCP_BENCH_SOURCE ?? project;
  if (!path.isAbsolute(requested)) throw new Error("OFFICEMCP_BENCH_SOURCE must be an absolute package directory.");
  const sourceRoot = fs.realpathSync(requested);
  for (const file of ["package.json", "src/cli.mjs", "src/contracts/registry.mjs"])
    if (!fs.statSync(path.join(sourceRoot, file), { throwIfNoEntry: false })?.isFile())
      throw new Error(`Benchmark source is missing ${file}.`);
  const packageJson = JSON.parse(fs.readFileSync(path.join(sourceRoot, "package.json"), "utf8"));
  if (packageJson.name !== "officemcp-local-preview" || packageJson.type !== "module")
    throw new Error("Benchmark source must be an OfficeMCP ESM package.");
  const dependenciesRoot = fs.realpathSync(path.join(sourceRoot, "node_modules"));
  if (!fs.statSync(dependenciesRoot).isDirectory()) throw new Error("Benchmark source dependencies are missing.");
  const label = env.OFFICEMCP_BENCH_LABEL ?? `run-${Date.now()}-${randomUUID().slice(0, 8)}`;
  if (!/^[a-z0-9][a-z0-9_-]{0,63}$/.test(label))
    throw new Error("OFFICEMCP_BENCH_LABEL must contain 1-64 lowercase letters, digits, underscores or hyphens and start with a letter or digit.");
  return { sourceRoot, dependenciesRoot, label };
}

export function benchmarkReportPath(project, basename, label) {
  if (!/^[a-z0-9][a-z0-9_-]{0,127}$/.test(basename) || !/^[a-z0-9][a-z0-9_-]{0,63}$/.test(label))
    throw new Error("Invalid benchmark report name or label.");
  return path.join(project, "test_doc", "officemcp-results", `${basename}-${label}.json`);
}

export function writeBenchmarkReport(file, value) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(value, null, 2) + "\n", { flag: "wx" });
}
