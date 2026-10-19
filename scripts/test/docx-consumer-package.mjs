/** Fresh npm dependency consumer, then complete jobs through the installed CLI. */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { hash, productIdentity, sdkIdentity } from './docx-product-evidence.mjs';
import { runWorkflows } from './docx-product-workflows.mjs';
import { runExpandedWorkflows } from './docx-expanded-workflows.mjs';
const project = fileURLToPath(new URL('../../', import.meta.url));
const output = path.resolve(process.env.OFFICEMCP_DOCX_OUTPUT ?? path.join(project, 'artifacts/docx-consumer', String(Date.now())));
fs.mkdirSync(output, { recursive: true });
const root = fs.mkdtempSync(path.join(os.tmpdir(), 'officemcp-docx-consumer-'));
const cache = process.env.npm_config_cache ?? path.join(os.tmpdir(), 'officemcp-npm-cache');
const report = { passed: false, platform: process.platform + '-' + process.arch, consumerRoot: root,
  installation: 'Fresh npm dependency installation; no overrides; lifecycle scripts disabled; offline by default.', sourceIdentities: {} };
const npm = process.platform === 'win32' ? process.execPath : 'npm';
const prefix = process.platform === 'win32' ? [process.env.npm_execpath ?? path.join(path.dirname(process.execPath), 'node_modules/npm/bin/npm-cli.js')] : [];
function run(args, cwd, log) {
  try { fs.writeFileSync(path.join(output, log), execFileSync(npm, [...prefix, ...args], { cwd, encoding: 'utf8', timeout: 180000, maxBuffer: 32 * 1024 * 1024 })); }
  catch (error) { fs.writeFileSync(path.join(output, log), String(error.stdout ?? '') + String(error.stderr ?? '')); throw error; }
}
try {
  let tarball = process.env.OFFICEMCP_DOCX_TARBALL;
  if (!tarball) {
    run(['pack', '--json', '--pack-destination', output, '--cache', cache], project, 'pack.json');
    tarball = path.join(output, JSON.parse(fs.readFileSync(path.join(output, 'pack.json')))[0].filename);
  }
  tarball = path.resolve(tarball); report.tarball = tarball; report.tarballHash = hash(fs.readFileSync(tarball));
  fs.writeFileSync(path.join(root, 'package.json'), JSON.stringify({ name: 'independent-docx-consumer', private: true, type: 'module' }));
  run(['install', tarball, '--ignore-scripts', '--omit=dev', '--no-audit', '--no-fund', '--cache', cache,
    ...(process.env.OFFICEMCP_DOCX_ONLINE === '1' ? [] : ['--offline'])], root, 'install.log');
  const sourceRoot = path.join(root, 'node_modules/officemcp-local-preview'); report.installedPackage = sourceRoot;
  report.version = JSON.parse(fs.readFileSync(path.join(sourceRoot, 'package.json'))).version;
  report.productIdentity = productIdentity(sourceRoot);
  report.sdk = sdkIdentity(sourceRoot);
  assert.equal(report.sdk.version, JSON.parse(fs.readFileSync(path.join(sourceRoot, 'package.json'))).dependencies['@superdoc/sdk']);
  if (!process.env.OFFICEMCP_DOCX_TARBALL) assert.deepEqual(report.sdk, sdkIdentity(project), 'Installed SDK/native executable differs from selected source');
  if (!process.env.OFFICEMCP_DOCX_TARBALL) assert.equal(report.productIdentity, productIdentity(project), 'Installed production source/contracts/core differ from selected source');
  for (const file of ['package.json', 'src/cli.mjs', 'contracts/full-tools.json', 'contracts/dependencies.json']) {
    report.sourceIdentities[file] = hash(fs.readFileSync(path.join(sourceRoot, file)));
    if (!process.env.OFFICEMCP_DOCX_TARBALL) assert.equal(report.sourceIdentities[file], hash(fs.readFileSync(path.join(project, file))));
  }
  const workflows = await runWorkflows({ sourceRoot, output: path.join(output, 'installed-workflows') });
  report.workflows = path.join(output, 'installed-workflows/workflows.json');
  const expanded = await runExpandedWorkflows({ sourceRoot, output: path.join(output, 'installed-expanded'), word: process.env.OFFICEMCP_WORD, wordScratch: process.env.OFFICEMCP_WORD_SCRATCH });
  report.expandedWorkflows = path.join(output, 'installed-expanded/expanded-workflows.json'); report.passed = workflows.passed && expanded.passed;
} catch (error) { report.error = String(error.stack ?? error); }
finally {
  report.consumerRetained = process.env.OFFICEMCP_DOCX_KEEP_CONSUMER === '1';
  if (!report.consumerRetained) fs.rmSync(root, { recursive: true, force: true });
  fs.writeFileSync(path.join(output, 'consumer.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify({ passed: report.passed, report: path.join(output, 'consumer.json'), error: report.error }));
  process.exitCode = report.passed ? 0 : 1;
}
