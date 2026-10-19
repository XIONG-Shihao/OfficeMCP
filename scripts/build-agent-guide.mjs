/** Generates the agent guide for OfficeMCP's DOCX tools from its own contracts.
 *   node scripts/build-agent-guide.mjs           write agent/OfficeMCP_docx.md
 *   node scripts/build-agent-guide.mjs --check   fail if the committed guide is stale */
import { readFileSync, writeFileSync } from 'node:fs'
import { pathToFileURL } from 'node:url'
import { z } from 'zod'
import {operationRegistry,VERSION,DOCX_AGENT_CONTRACT,DOCX_ACTION_REGISTRY,DOCX_ACTION_FAMILIES} from '../src/authoring.mjs'
export const guidePath = new URL('../agent/OfficeMCP_docx.md', import.meta.url)
const jsonSchema = schema => z.toJSONSchema(schema, { target: 'draft-2020-12', io: 'input' })
const block = text => '```ts\n' + text + '\n```'
const comment = text => String(text).replaceAll('`', '').split(/\r?\n/).flatMap(line => {
  const words = line.split(/\s+/), lines = []; let current = '//'
  for (const word of words) {
    if (current.length + word.length > 96) { lines.push(current); current = '//' }
    current += ' ' + word
  }
  return [...lines, current]
}).join('\n')
function fingerprint(schema) {
  return JSON.stringify(schema, (key, value) => ['description', '$schema'].includes(key) ? undefined : value)
}
function rules(schema) {
  const terms = []
  if (schema.type === 'integer') terms.push('integer')
  if (schema.default !== undefined) terms.push(`default ${JSON.stringify(schema.default)}`)
  for (const [key, label] of [['minimum', 'min'], ['maximum', 'max'], ['minLength', 'min length'], ['maxLength', 'max length'], ['minItems', 'min items'], ['maxItems', 'max items'], ['pattern', 'pattern']])
    if (schema[key] !== undefined) terms.push(`${label}: ${schema[key]}`)
  if (schema.items) {
    const itemRules = rules({ ...schema.items, description: undefined })
    if (itemRules) terms.push(`each item: ${itemRules}`)
  }
  for (const [key, label] of [['propertyNames', 'keys'], ['additionalProperties', 'values']]) {
    if (schema[key] && typeof schema[key] === 'object') {
      const constraints = rules(schema[key])
      if (constraints) terms.push(`${label}: ${constraints}`)
    }
  }
  if (schema.description) terms.push(schema.description)
  return terms.join('; ')
}
function typeRenderer(aliases) {
  function type(schema, indent = '', expand = false) {
    const alias = aliases.get(fingerprint(schema))
    if (alias && !expand) return alias
    if (schema.const !== undefined) return JSON.stringify(schema.const)
    if (schema.enum) return schema.enum.map(value => JSON.stringify(value)).join(' | ')
    if (schema.anyOf || schema.oneOf) return (schema.anyOf ?? schema.oneOf).map(s => type(s, indent)).join(' | ')
    if (schema.$ref === '#/$defs/__schema0') return 'JsonValue'
    if (schema.$ref) throw Error(`Unhandled documentation schema reference: ${schema.$ref}`)
    if (schema.type === 'array') return `Array<${type(schema.items ?? {}, indent)}>`
    if (schema.type === 'object' || schema.properties) {
      const properties = Object.entries(schema.properties ?? {})
      if (!properties.length) return schema.additionalProperties === false ? 'Record<string, never>' : `Record<string, ${schema.additionalProperties && typeof schema.additionalProperties === 'object' ? type(schema.additionalProperties, indent) : 'unknown'}>`
      const fields = properties.map(([key, child]) => {
        const field = `${indent}  ${/^[a-zA-Z_$][\w$]*$/.test(key) ? key : JSON.stringify(key)}${schema.required?.includes(key) ? '' : '?'}: ${type(child, indent + '  ')};`
        const constraints = rules(child)
        if (!constraints) return field
        if (!field.includes('\n') && field.length + constraints.length < 115) return `${field} // ${constraints}`
        return `${comment(constraints).split('\n').map(line => indent + '  ' + line).join('\n')}\n${field}`
      })
      const compact = `{ ${fields.map(field => field.trim()).join(' ')} }`
      if (compact.length < 85 && !compact.includes('//') && !compact.includes('\n')) return compact
      return `{\n${fields.join('\n')}\n${indent}}`
    }
    return ({ integer: 'number', string: 'string', number: 'number', boolean: 'boolean', null: 'null' })[schema.type] ?? 'unknown'
  }
  return type
}
const toolNotes = {
  docx_layout: 'Optional configured Word/LibreOffice provider. Records layoutId for this exact revision. Page targeting requires unique literal text that resolves to one native range; render again after editing. Reference pixel comparison is per artifact, not a universal Word appearance guarantee.',
  docx_open: 'Open an allowed source; keep the returned view. Opening is not a complete read.',
  docx_create: 'from_content requires document and excludes path. from_template requires path and excludes document. Returns a private candidate; save to publish. Images must be allowed PNG/JPEG assets.',
  docx_inspect: 'Omit view for engine-free capability help; includeDomains then requires view. Request actionFamily, actions or workflows for focused schemas/examples. With view, also inspect document counts/inventory. documentInspected distinguishes the modes; inventory alone is not a complete read.',
  docx_read: 'Follow nextOffset and coverage.remainingReads for the required scope. Reuse readSnapshotId with the same view for complete-read paging. Items depend on representation/domain; not every item has editable text/target.',
  docx_query: 'require=exactlyOne refuses missing/ambiguous matches. Entity selectors need at least one filter; table selectors need contains or header. Page selectors require docx_layout and a unique literal phrase; stale, ambiguous and cross-page targets are refused.',
  docx_apply: 'Exactly one of action, steps or workflow. mode=apply requires operationId; preview does not. Action properties go inside arguments. selectorTarget, selectorTargets and select are mutually exclusive. Inline step selectors resolve against each step\'s current revision. Optional readback returns bounded body changes and fresh targets on the returned view; maxChars is the total text budget.',
  docx_validate: 'Requires an unpublished candidate. Reports semantic/package checks, not page layout. Save obtains validation automatically.',
  docx_compare: 'Semantic comparison against the opened source baseline; coverage states what was compared. No pixel comparison.',
  docx_image_get: 'Use an image target from docx_read({representation:"domain", domain:"images", ...}). Returns exact embedded PNG/JPEG bytes as native MCP image content alongside structuredContent. Generic query targets may lack the package locator.',
  docx_export: 'Requires configured LibreOffice. Publishes PDF, not DOCX; never certifies Microsoft Word layout. Retry uncertain publication with the same operationId.',
  docx_save: 'Validates the candidate, then publishes to a NEW .docx destination; no overwrite. Source stays unchanged. Use a configured destination rootId. Save completes the product output; a host workspace commit is a separate operation.',
  docx_close: 'retain keeps unpublished candidates; discard removes them. Source and published files remain. A closed handle cannot be reused.',
  docx_operation_get: 'Retrieve the original operationId after an uncertain write. Result shape depends on the recorded operation. Pending operations return state/recovery details; failed/expired operations return an error.',
}
const notes = {
  insert_paragraphs:'Provide exactly one of text or texts. Only the first inserted paragraph receives headingLevel. before/after placement requires a block target.',
  apply_style:'Provide exactly one of styleId, headingLevel, or likeText.',
  insert_image:'Provide exactly one of assetId or path. before/after placement needs an anchor target. The image path must be within an allowed project root.',
  replace_image:'Provide exactly one of assetId or path. The existing image target is required; resetSize defaults to false.',
  crop_image:'Provide at least one crop edge. Values are percentages, not points.',
  set_section_layout:'Provide at least one page-size, orientation, margin or break property. Numeric dimensions and margins are points.',
  format_table_cell:'Provide at least one cell formatting property, not just changeMode.',
  format_table_row:'Provide heightPt, allowBreakAcrossPages or repeatHeader. heightRule alone does not satisfy the request.',
  format_paragraph_advanced:'Provide indentation, tabStop, borders or shading.',
  patch_hyperlink:'Provide at least one of href, anchor, docLocation, tooltip, target or rel.',
  write_table_cells:'Use zero-based row/column coordinates and exact expectedText. Duplicate coordinates, merged cells and cells with multiple paragraphs are refused by the qualified writer.',
  accept_tracked_changes:'Some complex formatted/table-bearing ranges are refused. A refusal must not be represented as completed review.',
  reject_tracked_changes:'Some complex formatted/table-bearing ranges are refused. A refusal must not be represented as completed review.',
}
const exampleView={documentId:'doc_RETURNED',revision:'sha256:'+'0'.repeat(64)}
export const guideExamples = [
  {tool:'docx_apply',arguments:{view:exampleView,operationId:'keyed-price-edit',select:{type:'table',header:'Price USD',contains:'Support'},workflow:{type:'update_table_by_key',keyColumn:'Service',updates:[{key:'Support',column:'Price USD',expectedText:'125',text:'150'}]},readback:{maxItems:5,maxChars:2000}}},
  {tool:'docx_apply',arguments:{view:exampleView,operationId:'section-edit',workflow:{type:'replace_section',headingPath:['Methodology'],subsections:'preserve',expectedText:'Existing method.',paragraphs:['Revised method.']},readback:{maxItems:5,maxChars:2000}}},
  {tool:'docx_open',arguments:{path:'/absolute/project/contract.docx'}},
  {tool:'docx_inspect',arguments:{actions:['replace_text'],workflows:['update_table_by_key','replace_section']}},
  {tool:'docx_query',arguments:{view:exampleView,select:{type:'table',header:'Service',contains:'Support'},require:'exactlyOne'}},
  {tool:'docx_apply',arguments:{view:exampleView,operationId:'price-edit-unique',selectorTarget:{id:'target_RETURNED'},workflow:{type:'update_table',cells:[{row:1,column:1,expectedText:'125',text:'150'}]}}},
  {tool:'docx_save',arguments:{view:{...exampleView,candidateId:'candidate_RETURNED'},operationId:'price-save-unique',destination:{rootId:'output',relativePath:'contract-updated.docx'}}},
  {tool:'docx_apply',arguments:{view:exampleView,mode:'preview',workflow:{type:'fill_template',bindings:[{placeholder:'{{CLIENT}}',value:'Meridian Labs',expectedCount:1}]}}},
  {tool:'docx_read',arguments:{view:exampleView,representation:'domain',domain:'images',limit:10}},
  {tool:'docx_apply',arguments:{view:exampleView,operationId:'logo-resize-unique',action:'resize_image',selectorTarget:{id:'target_RETURNED'},arguments:{width:144,height:57.6,unit:'pt'}}},
]
export function renderDocxGuide() {
  const operations = operationRegistry('docx-agent')
  const inputs = Object.fromEntries([...operations].map(([name, op]) => [name, jsonSchema(op.input)]))
  const apply = inputs.docx_apply.properties
  const readOutput = operations.get('docx_read').output.properties
  const shared = {
    View: apply.view, Target: apply.selectorTarget, Story: apply.story,
    Select: inputs.docx_query.properties.select,
    Destination: inputs.docx_save.properties.destination,
    ContentDocument: inputs.docx_create.properties.document,
    ActionName: apply.action, ActionStep: apply.steps.items,
    Postconditions: apply.expect,
    Coverage: readOutput.coverage, ReadItem: readOutput.data.properties.items.items,
  }
  const workflows = apply.workflow.anyOf ?? apply.workflow.oneOf
  const aliases = new Map(Object.entries(shared).map(([name, schema]) => [fingerprint(schema), name]))
  aliases.set(fingerprint(apply.workflow), 'Workflow')
  aliases.set(fingerprint(readOutput.view), 'View')
  aliases.set(fingerprint(readOutput.data.properties.items.items.properties.target), 'Target')
  const type = typeRenderer(aliases)
  const declaration = (name, schema) => `type ${name} = ${type(schema, '', true)};`
  const signature = (name, schema, result) => `declare function ${name}(input: ${type(schema)}): ${result};`
  const S = { type: 'string' }, B = { type: 'boolean' }, O = { type: 'object' }
  const shape = properties => ({ type: 'object', properties })
  // Generic result profiles omit some runtime receipt fields; document these as optional.
  const supplements = {
    docx_open: { capabilities: { type: 'array', items: S } },
    docx_create: { created: B, published: B, sourceTemplateUnchanged: B, candidateId: S, candidateRevision: S, nextAction: S },
    docx_validate: { validationId: S, revision: S, profile: S },
    docx_image_get: { mediaType: S },
    docx_save: { candidateId: S, validationPassed: B },
    docx_close: { closed: B, disposition: { enum: ['retain', 'discard'] } },
    docx_operation_get: { state: S, recovery: S, created: B, published: B, sourceTemplateUnchanged: B, path: S, assetId: S, outputRevision: S, format: S },
  }
  const resultNames = {
    docx_layout: 'LayoutResult',
    docx_open: 'OpenResult', docx_create: 'CreateResult', docx_inspect: 'InspectResult',
    docx_read: 'ReadResult', docx_query: 'ReadResult', docx_apply: 'ApplyResult',
    docx_validate: 'ValidationResult', docx_compare: 'CompareResult', docx_image_get: 'ImageResult',
    docx_export: 'ExportResult', docx_save: 'SaveResult', docx_close: 'CloseResult', docx_operation_get: 'OperationResult',
  }
  const resultDeclaration = (name) => {
    const props = { ...operations.get(name).output.properties.data.properties, ...supplements[name] }
    if (name === 'docx_apply') props.steps = { type: 'array', items: shape({ action: S, resolvedTargets: { type: 'array', items: S }, verificationPassed: B }) }
    let result = `type ${resultNames[name]} = McpResult<${type(shape(props))}>`
    if (name === 'docx_open') result += ' & { document?: { id: string; format: "docx"; sourceRevision: string } }'
    if (name === 'docx_save') result += ' & { validation?: { checks?: string[]; notVerified?: string[]; profile?: string } }'
    return result + ';'
  }
  const chunks = [
    `# OfficeMCP DOCX guide\n\n<!-- Generated by OfficeMCP/scripts/build-agent-guide.mjs. Edit the generator/contracts, not this generated template. -->\n\nPackage ${VERSION}; contract ${DOCX_AGENT_CONTRACT}. ${operations.size} tools, ${Object.keys(DOCX_ACTION_REGISTRY).length} actions, ${workflows.length} composite workflows.`,
    [
      '## Definitions: tool, action and workflow',
      '',
      '**A tool is what the agent calls through MCP. An action or workflow specifies work inside a tool.** Only "tool" is a standard MCP concept; "action" and "workflow" here are OfficeMCP concepts.',
      '',
      '| Term | Definition | Example |',
      '|---|---|---|',
      '| **Tool** | A named MCP entry point with an input schema and a returned result. It can read, edit, validate or publish a document. | `docx_open`, `docx_read`, `docx_apply`, `docx_save` |',
      '| **Action** | A named, bounded editing operation with its own arguments and targeting rules, selected through `docx_apply.action`. Its implementation may use one or several SuperDoc API calls, or OfficeMCP document-handling code. | `replace_text`, `insert_paragraphs`, `resize_image` |',
      '| **Workflow** | A named task with predefined editing logic and checks, selected through `docx_apply.workflow`. It may coordinate several actions; its behavior includes the guards and verification needed for that task. | `fill_template`, `update_table`, `normalize_document_style` |',
      '',
      'The relationship is not a fixed "many APIs → one action → many actions → one tool" hierarchy. An action may need only one API call, and tools such as open/read/save do not collect editing actions. `docx_apply` is the tool that offers the action and workflow choices:',
      '',
      '```text',
      'Agent calls docx_apply with exactly one of:',
      '  action    → one named edit, with its arguments',
      '  steps     → an ordered list of actions supplied by the agent',
      '  workflow  → a predefined task with its own payload and checks',
      '```',
      '',
      'For example, replacing "30 days" with "45 days" in a selected clause uses the `replace_text` action. Filling client/date placeholders while checking their expected occurrence counts uses the `fill_template` workflow. Both are invoked through `docx_apply`; publishing the candidate still requires `docx_save`.',
      '',
      'Action signatures later in this guide describe the action payload; they do not expose extra MCP tools. Agents call the tool, not SuperDoc APIs directly.',
    ].join('\n'),
    '## Start here\n\nTool names in this guide are unprefixed. Use the namespace advertised by your MCP host; the OfficeMCP product contract stays the same across hosts. Read only the sections needed for the task. If disconnected, report the missing MCP connection.\n\nOpen/create → inspect/read/query → apply → verify → save to a new file. Use the returned view after each mutation. Prefer direct edits; use tracked mode when Word redlines are requested. Save validates automatically. Use a document viewer for visual review.\n\nThese are **TypeScript-style reference signatures**, not a JavaScript SDK. Each MCP call takes **one JSON object**; `input` is the parameter name, not an extra JSON key. `?` means optional; defaults and limits appear in `//` comments. Only send documented keys. Returned optional fields depend on the operation/domain; check status/error before using them.',
    '## Shared types and rules\n\n' + block([
      '// IDs/targets/snapshots come from the server; never fabricate them. Source View omits candidateId.',
      '// Targets and readSnapshotId are revision-bound. After an edit, use the new view and fresh targets.',
      '// Unique operationId per distinct write. On uncertainty, retry identical arguments/ID or get its receipt.',
      '// A corrected request after a known failure needs a new operationId. Timeout does not mean no write.',
      '// TARGET_CONFLICT / REVISION_CONFLICT: resolve current view/targets. HANDLE_CLOSED: reopen.',
      '// INVALID_ARGUMENT: follow argumentIssues/schema. ENGINE_UNAVAILABLE: report engine requirement.',
      '// OPERATION_PENDING / PUBLICATION_UNCERTAIN: recover the same operation before new work.',
      '// Complete window != complete document: follow coverage.remainingReads for the requested scope.',
      '// No renderer tools, page selectors, layout targets or visualEvidence arguments.',
      '// Textbox/nested-table/style-definition/numbering-definition edits and some tracked ranges are limited.',
      '// SDK/package validation does not establish exact Microsoft Word layout.',
      'type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue };',
      ...Object.entries(shared).map(([name, schema]) => {
        const guard = name === 'Select' ? '// Entity: at least one filter; table: contains or header required.\n' : name === 'Postconditions' ? '// At least one property; text occurrence assertions, not visual assertions.\n' : ''
        return guard + declaration(name, schema)
      }),
      '// structuredContent below; also inspect the MCP protocol isError flag. data is operation-specific.',
      '// Native image content is alongside structuredContent, not a field in this envelope.',
      `type McpResult<T> = {
  contractVersion: "${DOCX_AGENT_CONTRACT}";
  requestId: string;
  status: "completed" | "failed";
  warnings: string[];
  error?: Record<string, unknown>;
  operationId?: string;
  view?: View;
  coverage?: Coverage;
  data?: T;
};`,
    ].join('\n\n').replace(/(\/\/[^\n]*)\n\n(?=\/\/)/g, '$1\n')),
    '## Tools\n\n' + [...operations.keys()].map(name => `[${name}](#${name})`).join(' · '),
  ]
  for (const [name] of operations) {
    const result = ['docx_query', 'docx_apply'].includes(name) ? '' : '\n\n' + resultDeclaration(name)
    chunks.push(`## ${name}\n\n` + block(comment(toolNotes[name]) + '\n' + signature(name, inputs[name], resultNames[name]) + result))
  }
  chunks.push('## Shared action result\n\nAll actions use `docx_apply`; the action signatures below describe only its `arguments` object. They are not separate MCP tools. Supply `view`, `operationId` and any `selectorTarget`/`selectorTargets`/`select` at the outer level.\n\n' + block([
    '// mode=apply returns a new candidate view; still unpublished until docx_save.',
    '// mode=preview returns the original view, preview/wouldChange, and no durable candidate.',
    '// Receipts are compact: use bounded read/query for content not included in the result.',
    '// Workflow results may also include workflow, workflowVerification, postconditions and semanticDelta.',
    '// readback is optional, body-scoped and bounded by maxItems/maxChars; truncated means more content exists.',
    '// Readback targets come from the saved/reopened candidate. Deleted items and previews have no targets.',
    '// Targets address whole blocks/tables, not matched text spans; re-query to format only a substring.',
    '// A tableCell readback target addresses its parent table; retain row/column for a cell action.',
    '// Readback is not complete document/formatting/review/layout coverage; follow notChecked and read explicitly.',
    '// Action targets: "*-optional" may be omitted only for a qualified placement/scope.',
    '// Coordinates are zero-based; tableOrdinal is one-based. Points are used where specified.',
    resultDeclaration('docx_apply'),
  ].join('\n')))
  chunks.push('## Composite workflows\n\nPass a `Workflow` as `docx_apply.workflow` (not `action` or `arguments`); returns `ApplyResult`.\n\n' + block([
    '// update_table_by_key: requires one table selectorTarget or inline table select; first row is header.',
    '// Header/key matching is exact (caseSensitive defaults true); duplicate/missing matches are refused.',
    '// expectedText guards each cell. No duplicate cell writes, merged/nested tables or complex target cells.',
    '// replace_section: exact top-level body headingPath; selected heading is retained; direct edits only.',
    '// subsections=preserve stops before the first subsection; replace extends to the next peer/ancestor heading.',
    '// expectedText joins the selected body paragraphs/headings with newline, excluding the retained heading.',
    '// Replaced range: at most 100 paragraph/heading blocks with plain text inlines; no tables/lists/controls.',
    '// This version refuses documents containing comments, tracked changes, bookmarks, fields or cross-references.',
    '// The replacement is plain paragraphs; outside-section content is checked before accepting the candidate.',
    '// fill_template: check each placeholder\'s expectedCount and verify replacements.',
    '// update_table: requires an opaque table selectorTarget; guarded cells and caption are one transaction.',
    '// normalize_document_style: fontFamily and/or fontSize required.',
    '// Explicit expect adds checks; it cannot override the workflow\'s required occurrence guards.',
    '// Preview with mode="preview"; applying requires operationId.',
    declaration('Workflow', apply.workflow),
  ].join('\n')))
  chunks.push('## Action family index\n\n' + DOCX_ACTION_FAMILIES.map(f => `- [${f.id}](#family-${f.id}) (${f.actions.length}): ${f.actions.map(name => `[${name}](#${name})`).join(', ')}`).join('\n'))
  for (const family of DOCX_ACTION_FAMILIES) {
    chunks.push(`## Family ${family.id}\n\n${family.description}`)
    for (const name of family.actions) {
      const action = DOCX_ACTION_REGISTRY[name]
      chunks.push(`### ${name}\n\n` + block([
        comment(action.description),
        `// Target: ${action.target}; cardinality: ${action.cardinality}; atomicity: ${action.atomicity}.`,
        ...(notes[name] ? [comment(notes[name])] : []),
        signature(name, jsonSchema(action.schema), 'ApplyResult'),
      ].join('\n')))
    }
  }
  chunks.push('## Call examples\n\nSeparate MCP calls, not a runnable batch. Replace paths, views and targets with actual values; update the view after each edit. The outer `arguments` below belongs to the MCP request; inside `docx_apply`, `arguments` holds the selected action payload.\n\n' + guideExamples.map(e => '```json\n' + JSON.stringify({ name: e.tool, arguments: e.arguments }) + '\n```').join('\n\n'))
  return chunks.join('\n\n') + '\n'
}
export function writeDocxGuide({ check = false } = {}) {
  const expected = renderDocxGuide()
  if (check) { if (readFileSync(guidePath, 'utf8') !== expected) throw Error('agent/OfficeMCP_docx.md is stale. Run node scripts/build-agent-guide.mjs.') }
  else writeFileSync(guidePath, expected)
  return expected
}
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) writeDocxGuide({ check: process.argv.includes('--check') })
