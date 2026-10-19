import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { manifest, CONTRACT } from '../src/contracts/registry.mjs';
import { DOCX_AGENT_CONTRACT } from '../src/formats/docx/docx-agent-contract.mjs';

// Preserve the original 0.12/docx/3 hashes. The small schema snapshot makes each
// intentional breaking removal auditable instead of approving replacement hashes.
const baseline = JSON.parse(fs.readFileSync(new URL('./fixtures/operation-contracts.json', import.meta.url)));
const beforeRemoval = JSON.parse(fs.readFileSync(new URL('./fixtures/render-removal-0.12-docx3.json', import.meta.url)));
const addedTools = new Set(['xlsx_query', 'xlsx_apply', 'docx_layout']);
const removedTools = new Set(['office_render', 'docx_render', 'pdf_render', 'pptx_render']);
const wire = ({ inputSchema, outputSchema, annotations, _meta }) => ({ inputSchema, outputSchema, annotations, _meta });
const hash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
function oldVersions(contract) {
  contract = structuredClone(contract);
  const metadata = contract._meta['io.officemcp/operation'];
  metadata.resultContract = metadata.resultContract === DOCX_AGENT_CONTRACT ? 'docx/3' : '0.12';
  if (contract.outputSchema.properties.contractVersion.const === DOCX_AGENT_CONTRACT)
    contract.outputSchema.properties.contractVersion.const = 'docx/3';
  return contract;
}
function beforePreset(name, preset) {
  const contract = structuredClone(beforeRemoval[name]);
  if (preset === 'full') contract._meta['io.officemcp/operation'].presets.push('full');
  return contract;
}
function afterRemoval(name, preset) {
  const contract = beforePreset(name, preset);
  const metadata = contract._meta['io.officemcp/operation'];
  metadata.resultContract = DOCX_AGENT_CONTRACT;
  contract.outputSchema.properties.contractVersion.const = DOCX_AGENT_CONTRACT;
  const data = contract.outputSchema.properties.data.properties;
  if (['docx_read', 'docx_query'].includes(name))
    for (const field of ['layoutId', 'page', 'pageCount', 'mapping', 'pixelVerified']) delete data[field];
  if (['docx_query', 'docx_apply'].includes(name)) {
    // Inline action selectors share the same public selector union.
    const removePageSelector = value => {
      if (!value || typeof value !== 'object') return;
      if (Array.isArray(value.oneOf)) value.oneOf = value.oneOf.filter(option => option.properties?.type?.const !== 'page');
      for (const child of Object.values(value)) removePageSelector(child);
    };
    removePageSelector(contract.inputSchema);
  }
  if (['docx_validate', 'docx_compare'].includes(name)) {
    delete contract.inputSchema.properties.visualEvidence;
    delete data.visualEvidence;
    metadata.supportedEngines = ['superdoc-v2'];
  }
  if (name === 'docx_image_get')
      for (const field of ['page', 'pageCount', 'layoutId', 'renderer', 'anchors']) delete data[field];
  // docx/5 intentionally removes comment writers; comment reads remain intact.
  const removeCommentWrites = schema => {
    if (!schema || typeof schema !== 'object') return;
    if (Array.isArray(schema.enum)) schema.enum = schema.enum.filter(value => !['add_comments', 'resolve_comments', 'reply_to_comment'].includes(value));
    for (const value of Object.values(schema)) removeCommentWrites(value);
  };
  removeCommentWrites(contract);
  return contract;
}
function beforeAdditiveDocxWorkflows(name, contract) {
  if (name !== 'docx_apply') return contract;
  contract = structuredClone(contract);
  // These additions predate XLSX. Keep the frozen prior-version fixture intact
  // and compare every other field. Their behavior has dedicated workflow tests.
  const input = contract.inputSchema, data = contract.outputSchema.properties.data;
  assert(!input.required?.includes('readback'));
  assert(!data.required?.includes('readback'));
  assert(!data.required?.includes('workflowVerification'));
  delete input.properties.readback;
  delete data.properties.readback;
  delete data.properties.workflowVerification;
  const variants = input.properties.workflow.oneOf;
  const additions = new Set(['update_table_by_key', 'replace_section']);
  assert.deepEqual(variants.filter(item => additions.has(item.properties.type.const)).map(item => item.properties.type.const).sort(), [...additions].sort());
  input.properties.workflow.oneOf = variants.filter(item => !additions.has(item.properties.type.const));
  return contract;
}
// The frozen fixtures remain unchanged. Normalize only the declared docx/6
// additions; execution and refusal behavior are covered by the expansion tests.
function beforeCapabilityExpansion(name, contract) {
  const result = structuredClone(contract);
  const actions = new Set(['replace_textbox_text', 'replace_equation_text', 'set_chart_title', 'set_chart_data', 'replace_table_cell_text', 'replace_section_content']);
  function visit(schema) {
    if (!schema || typeof schema !== 'object') return;
    if (Array.isArray(schema.enum)) schema.enum = schema.enum.filter(v => !actions.has(v) && !['objects', 'table_cells', 'section_content'].includes(v));
    if (Array.isArray(schema.oneOf)) schema.oneOf = schema.oneOf.filter(v => v.properties?.type?.const !== 'page');
    if (schema.properties?.scope?.enum && JSON.stringify(schema.properties.scope.enum) === JSON.stringify(['body'])) schema.properties.scope = { const: 'body' };
    if (schema.properties?.targetKind?.enum) schema.properties.targetKind.enum = schema.properties.targetKind.enum.filter(v => v !== 'object');
    for (const value of Object.values(schema)) visit(value);
  }
  visit(result);
  if (['docx_read', 'docx_query', 'docx_apply', 'docx_validate', 'docx_save'].includes(name)) result._meta['io.officemcp/operation'].supportedEngines = result._meta['io.officemcp/operation'].supportedEngines.filter(e => e !== 'guarded-ooxml');
  if (name === 'docx_read') { delete result.inputSchema.properties.maxChars; delete result.inputSchema.properties.section; }
  if (name === 'docx_inspect') delete result.outputSchema.properties.data.properties.workflowLimits;
  return result;
}
for (const [preset, expected] of Object.entries(baseline))
  test(`core 0.13/docx/6 declares its additions while preserving ${preset} contracts`, () => {
    assert.equal(CONTRACT, '0.13');
    assert.equal(DOCX_AGENT_CONTRACT, 'docx/6');
    const actual = Object.fromEntries(manifest(preset).filter(tool => !addedTools.has(tool.name)).map(tool => {
      assert(!removedTools.has(tool.name));
      const contract = beforeCapabilityExpansion(tool.name, wire(tool));
      if (['docx-agent', 'full'].includes(preset) && beforeRemoval[tool.name]) {
        assert.deepEqual(beforeAdditiveDocxWorkflows(tool.name, contract), afterRemoval(tool.name, preset), `Unexpected change to ${preset}/${tool.name}`);
        return [tool.name, hash(beforePreset(tool.name, preset))];
      }
      const original = oldVersions(contract);
      if (['office_read', 'docx_read_details'].includes(tool.name)) {
        // Existing comment reading is additive on legacy detail domains.
        const removeReadDomainAddition = schema => {
          if (!schema || typeof schema !== 'object') return;
          if (Array.isArray(schema.enum) && schema.enum.includes('paragraphs') && schema.enum.includes('capabilities')) schema.enum = schema.enum.filter(v => v !== 'comments');
          for (const value of Object.values(schema)) removeReadDomainAddition(value);
        };
        removeReadDomainAddition(original.inputSchema);
      }
      if (['office_save', 'office_commit'].includes(tool.name)) {
        // Portable artifact publication adds one optional output descriptor.
        // Assert its exact shape, then compare all older fields to frozen hashes.
        const data = original.outputSchema.properties.data;
        assert(!data.required?.includes('artifactDescriptor'));
        assert.deepEqual(data.properties.artifactDescriptor, { type: 'string' });
        delete data.properties.artifactDescriptor;
      }
      if (tool.name === 'office_create') {
        const variants = original.inputSchema.properties.document.oneOf;
        assert.equal(variants.filter(item => item.properties?.format?.const === 'csv').length, 1);
        // CSV creation is an additive format. Preserve every prior format schema
        // and the original frozen compatibility hashes.
        original.inputSchema.properties.document.oneOf = variants.filter(item => item.properties?.format?.const !== 'csv');
      }
      // Prior additive family discovery, separately tested by discovery.test.
      if (tool.name === 'docx_inspect' && original.outputSchema.properties.contractVersion.const === 'docx/3') {
        // Restore only the three intentionally removed writer enum values for the frozen docx/3 hash.
        const actions = original.inputSchema.properties.actions.items.enum;
        assert.equal(actions.length, 73);
        assert(!actions.includes('add_comments'));
        actions.splice(actions.indexOf('rewrite_block'), 0, 'add_comments', 'resolve_comments', 'reply_to_comment');
        delete original.inputSchema.properties.actionFamily;
        delete original.outputSchema.properties.data.properties.actionFamilies;
        delete original.outputSchema.properties.data.properties.selectedActionFamily;
        // Additive workflow discovery can run without opening a document. The
        // frozen contract required view and had no workflow discovery fields.
        delete original.inputSchema.properties.workflows;
        for (const field of ['availableWorkflows', 'workflowDescriptors', 'usage', 'documentInspected'])
          delete original.outputSchema.properties.data.properties[field];
        const { additionalProperties, ...input } = original.inputSchema;
        original.inputSchema = { ...input, required: ['view'], additionalProperties };
      }
      if (tool.name === 'docx_operation_get') {
        // Receipts expose the same optional readback as docx_apply.
        delete original.outputSchema.properties.data.properties.readback;
        delete original.outputSchema.properties.data.properties.workflowVerification;
      }
      return [tool.name, hash(original)];
    }));
    assert.deepEqual(actual, Object.fromEntries(Object.entries(expected).filter(([name]) => !removedTools.has(name))));
  });

test('shared workbook tools are additive and available on common workflow surfaces', () => {
  const workbookAdditions = new Set(['xlsx_query', 'xlsx_apply']);
  for (const preset of ['all', 'recommended', 'full']) {
    const tools = manifest(preset);
    assert.deepEqual(tools.filter(tool => workbookAdditions.has(tool.name)).map(tool => tool.name).sort(), [...workbookAdditions].sort());
  }
  assert(!manifest('docx-agent').some(tool => workbookAdditions.has(tool.name)));
  for (const preset of ['docx-agent', 'full']) assert(manifest(preset).some(tool => tool.name === 'docx_layout'));
});
