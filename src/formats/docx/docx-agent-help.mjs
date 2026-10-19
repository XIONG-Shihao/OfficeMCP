import { z } from "zod";
import { docxActionDiscovery } from "./docx-agent-actions.mjs";
import { compositeWorkflow } from "./docx-agent-contract.mjs";
const examples = {
  fill_template: { type: "fill_template", bindings: [{ placeholder: "{{CLIENT}}", value: "Meridian Labs", expectedCount: 1 }] },
  update_table: { type: "update_table", cells: [{ row: 1, column: 1, expectedText: "125", text: "150" }] },
  normalize_document_style: { type: "normalize_document_style", fontFamily: "Arial", fontSize: 11 },
  update_table_by_key: { type: "update_table_by_key", keyColumn: "Service", updates: [{ key: "Support", column: "Price USD", expectedText: "125", text: "150" }] },
  replace_section: { type: "replace_section", headingPath: ["Methodology"], subsections: "preserve", expectedText: "Existing method.", paragraphs: ["Revised method."] },
};
const notes = {
  fill_template: "Checks expected occurrences before editing and replacements afterward.",
  update_table: "Requires a table selectorTarget. Guarded zero-based cells; optional caption in the same transaction.",
  normalize_document_style: "Provide fontFamily and/or fontSize.",
  update_table_by_key: "Requires a table selectorTarget or inline table select. First row is the header; exact unique header/key matching, expectedText required. Uses guarded SDK text replacement for unique plain value cells, preserving formatting and unrelated merges/nested content. Selected key/value cells must be unmerged. For merged owners or nested tables, read domain table_cells, select its explicit target, and use replace_table_cell_text. Never guess cell coordinates after refusal.",
  replace_section: "Body top-level headingPath, exact match. Read domain section_content with section:{headingPath,subsections} first. expectedText joins the returned item texts with newline, excluding the retained heading. Provide paragraphs or blocks (paragraph/heading/list/table/retain). Tables use tabs between cells and newlines between rows. SDK handles ordinary tables/lists. Explicit retain indices keep objects/protected blocks through the guarded package adapter. Object/review/reference deletion is refused. Outside content is verified.",
};
export function docxCapabilityHelp(args = {}) {
  const selected = args.workflows ?? [];
  return {
    ...docxActionDiscovery(args),
    availableWorkflows: Object.keys(examples),
    workflowLimits: {
      update_table_by_key: { target: "top-level body table", header: "unique unmerged first-row names", key: "one exact unique row key in an unmerged key cell", value: "unmerged one-paragraph value cell without protected objects", unrelatedMerges: "preserved", mutationMethod: "guarded SDK range replacement", refused: ["merged key/value cells", "merged headers", "nested selected table", "protected or multi-paragraph value cell", "duplicate keys/headers", "stale values"], nextAction: "Use this workflow first for supported cells. Do not guess coordinates after a refusal; narrow the target or read the stated limits." },
      tableCells: { readDomain: 'table_cells', target: 'explicit physical cell or recorded vertical-merge owner', action: 'replace_table_cell_text', expectedText: 'own paragraphs joined with newline; excludes nested tables', find: 'unique literal in one own paragraph; paragraphIndex disambiguates repeated paragraphs; empty find inserts into one empty paragraph', engine: 'SDK range when available, otherwise guarded OOXML', refused: ['inconsistent merge ownership', 'protected/object-bearing selected paragraphs', 'implicit structural deletion'] },
      replace_section: { target: "one exact body headingPath", read: 'section_content with section scope returns retain indices', outsideReviewContent: "preserved", protectedInsideRange: 'must be explicitly retained', replacement: "paragraphs or structured paragraph/heading/list/table/retain blocks", maxOldBlocks: 100 },
      charts: { title: 'existing direct rich-text title', data: 'one unique series; existing point count; literal caches or simple local A1 references to an embedded XLSX', consistency: 'cache and literal worksheet updated together', refused: ['external or formula workbooks', 'shared chart workbook/ranges', 'pivot charts', 'computed/missing titles', 'adding/removing points'] },
      page: { provider: "optional configured layout provider", target: "layoutId + page number + unique literal text", staleAfterEdit: true, refused: ["ambiguous text", "cross-page ranges", "opaque objects"], appearance: "measured comparison against an explicit Word PDF reference" },
    },
    workflowDescriptors: Object.fromEntries(compositeWorkflow.options
      .filter(schema => selected.includes(schema.shape.type.value))
      .map(schema => {
        const name = schema.shape.type.value;
        return [name, { name, description: notes[name], argumentSchema: z.toJSONSchema(schema, { io: "input" }),
          example: { tool: "docx_apply", arguments: {
            view: { documentId: "RETURNED_DOCUMENT_ID", revision: "sha256:" + "0".repeat(64) },
            mode: "preview", workflow: examples[name], readback: { maxItems: 5, maxChars: 2000 },
            ...(["update_table", "update_table_by_key"].includes(name) ? { selectorTarget: { id: "RETURNED_TABLE_TARGET" } } : {}),
          } },
        }];
      })),
    usage: {
      pages: "Call docx_layout first; query a recorded page using a precise literal phrase, then edit its returned semantic target. Render again after editing. A page number alone is never a mutation target.",
      comments: "Existing comments may be read and queried. OfficeMCP does not create, edit, reply to, delete, resolve or reopen comments.",
      toolPrefix: "Use the tool name and namespace advertised by your MCP connection.",
      examples: "Example IDs/revisions are placeholders. Replace them with returned values. Apply mode requires a unique operationId.",
      action: "docx_apply({view, operationId, action, arguments, selectorTarget?, readback?}); view is always nested.",
      receipt: "Use the returned view and readback targets. Save publishes a new file; previews have no reusable targets.",
    },
  };
}
