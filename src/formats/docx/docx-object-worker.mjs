import fs from 'node:fs';
import { openFilePackage } from '../opc/file-package.mjs';
import { rewriteFilePackage } from '../opc/file-package-writer.mjs';
import { inspectDocxPackage } from './docx-package-inventory.mjs';
import { paged } from './docx-canonical.mjs';
import { prepareObjectChanges } from './docx-object-edits.mjs';
import { fileRevision } from '../../infrastructure/files.mjs';
import { requireThat } from '../../infrastructure/errors.mjs';
import { parseXml } from '../opc/ooxml.mjs';
import { packageCells } from './docx-package-targets.mjs';
import { sectionInventory } from './docx-section-adapter.mjs';
const ENGINE = 'OfficeMCP guarded OOXML object adapter v1';

function sameBytes(left, right) { return Buffer.from(left).equals(Buffer.from(right)); }
function verifyParts(source, output, replacements) {
  requireThat(JSON.stringify(Object.keys(source.parts).sort()) === JSON.stringify(Object.keys(output.parts).sort()), 'PRESERVATION_FAILED', 'Object editing changed package membership.');
  for (const name of Object.keys(source.parts)) {
    const expected = replacements.get(name) ?? source.parts[name];
    requireThat(sameBytes(expected, output.parts[name]), 'PRESERVATION_FAILED', 'An unapproved package member differs: ' + name);
  }
}
export async function runObjectWorker(task) {
  if (["docx_opc_agent_read", "docx_opc_agent_query"].includes(task.kind)) {
    const domain = task.kind === 'docx_opc_agent_query' ? task.args.select.domain : task.args.domain;
    let items;
    if (['table_cells', 'section_content'].includes(domain)) {
      const source = openFilePackage(task.input, task.maxExpandedBytes, task.maxFileBytes);
      try { const document = parseXml(Buffer.from(source.parts['word/document.xml']).toString('utf8'));
        const parsed = new Map([['word/document.xml', document]]);
        if (domain === 'section_content' && source.parts['word/styles.xml']) parsed.set('word/styles.xml', parseXml(Buffer.from(source.parts['word/styles.xml']).toString('utf8')));
        items = domain === 'table_cells' ? packageCells(document).map(item => ({ ...item, contentSource: 'read-only-ooxml-package' })) : sectionInventory(parsed, task.args.section);
      } finally { source.close(); }
    } else items = inspectDocxPackage(task.input, task.maxExpandedBytes).complexObjects;
    if (task.kind === "docx_opc_agent_query") {
      const select = task.args.select, normalize = value => select.caseSensitive ? String(value ?? "") : String(value ?? "").toLocaleLowerCase();
      requireThat(['tag', 'author', 'changeType'].every(key => select[key] === undefined), 'UNSUPPORTED_FEATURE', 'Package object/cell/block selectors support id, name and text; tag, author and changeType are not applicable.');
      items = items.filter(item => ["id", "name", "text"].every(key => select[key] === undefined || normalize(item[key]) === normalize(select[key])));
      requireThat(task.args.require !== "exactlyOne" || items.length === 1, "TARGET_CONFLICT", "Object selector must resolve exactly once.");
    }
    let remaining = task.args.maxChars ?? 16000;
    const result = paged(items, task.args, domain, { engine: "read-only OOXML package inventory", notInterpreted: ["embedded payloads", "chart formulas", "SmartArt layout", "mathematics"] });
    if (['table_cells', 'section_content'].includes(domain)) result.items = result.items.map(item => {
      const text = item.text.slice(0, remaining); remaining -= text.length;
      const details = { ...item.details, text, textLength: item.text.length, truncated: text.length !== item.text.length };
      if (details.paragraphs) details.paragraphs = details.paragraphs.map(p => { const text = p.text.slice(0, remaining); remaining -= text.length; return { ...p, text, textLength: p.text.length, truncated: text.length !== p.text.length }; });
      return { ...item, text, details };
    });
    return result;
  }
  if (task.kind === 'docx_opc_validate') {
    let revision = fileRevision(task.original, task.maxFileBytes);
    for (const operation of task.operations ?? []) {
      requireThat(operation.engineEvidence?.verification === 'docx-opc-objects-v1' && operation.engineEvidence.inputRevision === revision,
        'VALIDATION_FAILED', 'Object candidate evidence chain is incomplete.');
      revision = operation.engineEvidence.outputRevision;
    }
    requireThat(revision === fileRevision(task.input, task.maxFileBytes), 'VALIDATION_FAILED', 'Object candidate revision changed.');
    // Reconstruct only the explicitly declared changes from the original.
    const source = openFilePackage(task.original, task.maxExpandedBytes, task.maxFileBytes), output = openFilePackage(task.input, task.maxExpandedBytes, task.maxFileBytes);
    try {
      const changes = prepareObjectChanges(source, (task.operations ?? []).flatMap(operation => operation.steps));
      verifyParts(source, output, changes.parts);
      return { passed: true, revision, profile: 'docx.guarded-objects.v1', checks: ['revision-chain', 'declared-object-field-readback', 'all-unselected-members-byte-identical'], notVerified: ['visual-layout', 'arbitrary-object-semantics'] };
    } finally { source.close(); output.close(); }
  }
  const source = openFilePackage(task.input, task.maxExpandedBytes, task.maxFileBytes);
  try {
    const changes = prepareObjectChanges(source, task.args.steps);
    const steps = task.args.steps.map((step, index) => ({ action: step.action, arguments: step.arguments, targetLocator: step.targetLocator,
      resolvedTargets: [step.targetLocator.objectId], receipt: { status: 'ok', verificationPassed: true, selectedTargets: [step.targetLocator.objectId] } }));
    let remaining = task.args.readback?.maxChars ?? 0;
    const boundedReadback = changes.readback.map(item => {
      const values = {};
      for (const side of ["before", "after"]) {
        const text = item[side].text.slice(0, remaining); remaining -= text.length;
        values[side] = { text, truncated: text.length < item[side].text.length };
      }
      return { ...item, ...values };
    });
    const readback = task.args.readback ? { scope: 'objects', items: boundedReadback.slice(0, task.args.readback.maxItems), total: changes.readback.length,
      truncated: changes.readback.length > task.args.readback.maxItems, reusableTargets: task.kind !== 'docx_opc_agent_preview', notChecked: ['visual-layout'] } : undefined;
    if (task.kind === 'docx_opc_agent_preview') return { preview: true, applied: false, wouldChange: steps.length > 0, steps, ...(readback ? { readback: { ...readback, items: readback.items.map(({ nativeTarget, ...item }) => item) } } : {}) };
    const saved = rewriteFilePackage(task.input, task.output, changes.parts, task), output = openFilePackage(task.output, task.maxExpandedBytes, task.maxFileBytes);
    try { verifyParts(source, output, changes.parts); } finally { output.close(); }
    const receipt = { status: 'ok', intent: steps.length === 1 ? steps[0].action : 'mutation_plan', verificationPassed: true, steps };
    return { revision: saved.revision, engine: ENGINE, profile: 'docx.guarded-objects.v1', action: receipt.intent, receipt, steps, affectedTargets: steps.flatMap(step => step.resolvedTargets), ...(readback ? { readback } : {}),
      engineEvidence: { engine: ENGINE, verification: 'docx-opc-objects-v1', inputRevision: source.revision, outputRevision: saved.revision,
        changedParts: saved.changedParts, checks: ['exact-object-guards', 'declared-field-readback', 'all-unselected-members-byte-identical'], notVerified: ['visual-layout', 'arbitrary-object-semantics'] } };
  } catch (error) { if (task.output) fs.rmSync(task.output, { force: true }); throw error; }
  finally { source.close(); }
}
