import { z } from "zod";
import { DOCX_AGENT_ACTIONS, DOCX_ACTION_FAMILY_IDS, sectionBlocks } from "./docx-agent-actions.mjs";
import { documentView, creation } from "../../contracts/workflow.mjs";
import { docxAgentResultSchemas } from "./docx-agent-results.mjs";

const id = z.string().min(1).max(150);
const page = {
  offset: z.number().int().min(0).default(0),
  limit: z.number().int().min(1).max(100).default(30),
};
const jsonValue = z.json();
const argumentsObject = z
  .record(z.string(), jsonValue)
  .default({})
  .describe(
    "Arguments for the selected action. Request its descriptor with docx_inspect.actions when needed; unknown or incomplete action arguments fail instead of being guessed.",
  );

export const DOCX_AGENT_CONTRACT = "docx/6";

const storySelector = z.discriminatedUnion("type", [
  z.object({ type: z.literal("body") }).strict(),
  z
    .object({
      type: z.literal("header"),
      sectionId: id,
      variant: z.enum(["default", "first", "even"]).default("default"),
    })
    .strict(),
  z
    .object({
      type: z.literal("footer"),
      sectionId: id,
      variant: z.enum(["default", "first", "even"]).default("default"),
    })
    .strict(),
  z.object({ type: z.literal("footnote"), id }).strict(),
  z.object({ type: z.literal("endnote"), id }).strict(),
  z.object({ type: z.literal("textbox"), id }).strict(),
]);

const querySelect = z.discriminatedUnion("type", [
  z.object({ type: z.literal("page"), number: z.number().int().min(1).max(2000), layoutId: id, pattern: z.string().min(1).max(32000) }).strict(),
  z
    .object({
      type: z.literal("text"),
      pattern: z.string().min(1).max(32000),
      mode: z.enum(["contains", "regex"]).default("contains"),
      caseSensitive: z.boolean().default(true),
      wholeWord: z.boolean().default(false),
    })
    .strict(),
  z
    .object({
      type: z.literal("node"),
      nodeType: z.enum([
        "paragraph",
        "heading",
        "listItem",
        "table",
        "tableRow",
        "tableCell",
        "tableOfContents",
        "image",
        "sdt",
        "run",
        "bookmark",
        "comment",
        "hyperlink",
        "footnoteRef",
        "endnoteRef",
        "crossRef",
        "tab",
        "lineBreak",
      ]),
    })
    .strict(),
  z
    .object({
      type: z.literal("entity"),
      domain: z.enum([
        "content_controls",
        "bookmarks",
        "captions",
        "tracked_changes",
        "comments",
        "hyperlinks",
        "fields",
        "footnotes",
        "endnotes",
        "cross_references",
            "objects", "table_cells", "section_content",
      ]),
      id: z.string().min(1).max(1000).optional(),
      name: z.string().min(1).max(1000).optional(),
      text: z.string().min(1).max(32000).optional(),
      tag: z.string().min(1).max(1000).optional(),
      author: z.string().min(1).max(1000).optional(),
      changeType: z
        .enum(["insert", "delete", "replacement", "format"])
        .optional(),
      caseSensitive: z.boolean().default(true),
    })
    .strict()
    .refine(
      (value) =>
        [value.id, value.name, value.text, value.tag, value.author, value.changeType]
          .some((item) => item !== undefined),
      { message: "Provide at least one entity filter." },
    ),
  z
    .object({
      type: z.literal("headingPath"),
      path: z.array(z.string().min(1).max(1000)).min(1).max(20),
      caseSensitive: z.boolean().default(false),
    })
    .strict(),
  z
    .object({
      type: z.literal("table"),
      contains: z.string().min(1).max(32000).optional(),
      header: z.string().min(1).max(32000).optional(),
      caseSensitive: z.boolean().default(false),
    })
    .strict()
    .refine((value) => value.contains || value.header, {
      message: "Provide table contains or header.",
    }),
]);

const postconditions = z
  .object({
    beforeOccurrences: z.record(nonemptyString(), z.number().int().min(0).max(100000)).optional(),
    remainingOccurrences: z.record(nonemptyString(), z.number().int().min(0).max(100000)).optional(),
    changedOccurrences: z.number().int().min(0).max(100000).optional(),
  })
  .strict()
  .refine((value) => Object.keys(value).length > 0, {
    message: "Provide at least one postcondition.",
  });

function nonemptyString() {
  return z.string().min(1).max(32000);
}

const workflowTableCell = z
  .object({
    row: z.number().int().min(0),
    column: z.number().int().min(0),
    expectedText: z.string().max(32000),
    text: z.string().max(32000),
  })
  .strict();

export const compositeWorkflow = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("update_table_by_key"),
    keyColumn: nonemptyString(),
    updates: z.array(z.object({
      key: nonemptyString(), column: nonemptyString(),
      expectedText: z.string().max(32000), text: z.string().max(32000),
    }).strict()).min(1).max(100),
    caseSensitive: z.boolean().default(true),
  }).strict(),
  z.object({
    type: z.literal("replace_section"),
    headingPath: z.array(z.string().min(1).max(1000)).min(1).max(20),
    subsections: z.enum(["preserve", "replace"]),
    expectedText: z.string().max(32000).describe("Exact selected section body text, joined with newline; excludes the section heading. preserve stops before the first subsection; replace includes subsection headings/text."),
    paragraphs: z.array(z.string().min(1).max(32000)).min(1).max(100).optional(),
    blocks: sectionBlocks.optional(),
  }).strict().refine(a => Boolean(a.paragraphs) !== Boolean(a.blocks), { message: 'Provide paragraphs or structured blocks, not both.' }),
  z
    .object({
      type: z.literal("fill_template"),
      bindings: z
        .array(
          z
            .object({
              placeholder: nonemptyString(),
              value: z.string().max(32000),
              expectedCount: z.number().int().min(1).max(1000),
            })
            .strict(),
        )
        .min(1)
        .max(200),
      changeMode: z.enum(["direct", "tracked"]).default("direct"),
    })
    .strict(),
  z
    .object({
      type: z.literal("update_table"),
      cells: z.array(workflowTableCell).min(1).max(100),
      caption: z
        .object({
          position: z.enum(["above", "below"]).default("below"),
          label: nonemptyString().default("Table"),
          text: z.string().max(32000),
        })
        .strict()
        .optional(),
    })
    .strict(),
  z
    .object({
      type: z.literal("normalize_document_style"),
      fontFamily: z.string().min(1).max(100).optional(),
      fontSize: z.number().min(1).max(200).optional(),
      changeMode: z.enum(["direct", "tracked"]).default("direct"),
    })
    .strict()
    .refine((value) => value.fontFamily || value.fontSize, {
      message: "Provide fontFamily or fontSize.",
    }),
]);

const selectorTarget = z.object({ id }).strict();
const actionStep = z
  .object({
    action: z.enum(DOCX_AGENT_ACTIONS),
    arguments: argumentsObject,
    selectorTarget: selectorTarget.optional(),
    selectorTargets: z.array(selectorTarget).min(1).max(100).optional(),
    select: querySelect.optional().describe(
      "Inline selector resolved with require=exactlyOne against the step's current revision.",
    ),
    story: storySelector.optional(),
  })
  .strict();

const docxAgentDefinitions = {
  docx_open: {
    description: "Open one allowed DOCX source and return a revision-bound view. Opening inventories the file but does not claim that every story or visual page was read.",
    input: z.object({ path: z.string().min(1).max(4096) }).strict(),
    output: docxAgentResultSchemas.generic,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "synchronous",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_create: {
    description: "Create a DOCX candidate with headings, text, tables and images. The source remains private until docx_save. Use from_template to copy an allowed DOCX without changing it.",
    input: z.object({
      operationId: id,
      action: z.enum(["from_content", "from_template"]).default("from_content"),
      document: creation.options
        .find((option) => option.shape.format.value === "docx")
        .optional(),
      path: z.string().min(1).max(4096).optional(),
    }),
    output: docxAgentResultSchemas.generic,
    risk: "mutate",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_layout: {
    description: "Render the exact view with an optional configured layout provider. Return revision-bound pages/layoutId; query a page with unique literal text before editing. Optional operator-supplied Word PDF reference compares exact pixels for this revision, without claiming universal Word equivalence.",
    input: z.object({ view: documentView, provider: z.enum(["configured", "libreoffice", "word"]).default("configured"), page: z.number().int().min(1).max(2000).optional(),
      offset: z.number().int().min(0).default(0), limit: z.number().int().min(1).max(20).default(5), maxChars: z.number().int().min(1).max(16000).default(2000),
      reference: z.object({ path: z.string().min(1).max(4096), documentRevision: z.string().regex(/^sha256:[a-f0-9]{64}$/), producerVersion: z.string().min(1).max(100) }).strict().optional() }).strict(),
    output: docxAgentResultSchemas.generic, risk: "read", handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible", engines: ["optional-layout-provider", "mupdf"], presets: ["docx-agent"],
  },
  docx_inspect: {
    description: "Without view, discover actions/workflows and examples without opening a document. With view, also return document counts, outline, diagnostics and a reading plan. Request actions, actionFamily or workflows for focused schemas.",
    input: z
      .object({
        view: documentView.optional(),
        includeDomains: z
          .array(
            z.enum([
              "blocks",
              "lists",
              "tables",
              "comments",
              "trackedChanges",
              "sections",
              "headerFooters",
              "styles",
              "contentControls",
              "fields",
              "hyperlinks",
              "bookmarks",
              "permissionRanges",
              "images",
              "footnotes",
              "endnotes",
              "toc",
              "captions",
              "crossReferences",
            ]),
          )
          .max(16)
          .optional(),
        actionFamily: z.enum(DOCX_ACTION_FAMILY_IDS).optional().describe(
          "Select a family from actionFamilies to return only its action names and at most 20 exact argument schemas. An explicit actions list overrides the family descriptor selection.",
        ),
        actions: z
          .array(z.enum(DOCX_AGENT_ACTIONS))
          .max(20)
          .optional()
          .describe(
            "Return machine-readable descriptors only for these actions.",
          ),
        workflows: z.array(z.enum(compositeWorkflow.options.map(option => option.shape.type.value))).max(5).optional(),
      })
      .strict().refine(value => value.view || !value.includeDomains?.length, { message: "includeDomains requires a document view." }),
    output: docxAgentResultSchemas.inspect,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_read: {
    description: "Read a canonical DOCX representation. complete returns one bounded cross-domain snapshot for broad small-document review; linear returns ordered blocks; structured returns hierarchy; domain returns one typed SDK domain. Always follow coverage and remainingReads until required scopes are complete.",
    input: z
      .object({
        view: documentView,
        representation: z
          .enum([
            "inventory",
            "complete",
            "linear",
            "structured",
            "markdown",
            "text",
            "domain",
          ])
          .default("linear"),
        story: storySelector.optional(),
        domain: z
          .enum([
            "paragraphs",
            "tables",
            "sections",
            "headers_footers",
            "fields",
            "footnotes",
            "endnotes",
            "hyperlinks",
            "images",
            "tracked_changes",
            "comments",
            "content_controls",
            "lists",
            "bookmarks",
            "toc",
            "captions",
            "cross_references",
            "objects", "table_cells", "section_content",
          ])
          .optional(),
        maxChars: z.number().int().min(0).max(16000).optional().describe('Aggregate text budget for package table-cell and section-content reads.'),
        section: z.object({ headingPath: z.array(nonemptyString()).min(1).max(20), subsections: z.enum(['preserve', 'replace']).default('preserve') }).strict().optional().describe('For section_content only: return section-relative block indices for retain.'),
        includeAnnotations: z.boolean().default(true),
        readSnapshotId: id
          .optional()
          .describe(
            "Explicit revision-bound read snapshot returned by an earlier complete read. Reuse it with the next offset to avoid rebuilding every domain.",
          ),
        ...page,
      })
      .strict(),
    output: docxAgentResultSchemas.read,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2", "guarded-ooxml"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_query: {
    description: "Resolve deterministic, mutation-grade targets against the exact view. Use require=exactlyOne for a unique clause. Returned opaque targets can be passed to docx_apply arguments as selectorTarget.",
    input: z
      .object({
        view: documentView,
        select: querySelect,
        story: storySelector.optional(),
        withinTarget: selectorTarget.optional(),
        require: z.enum(["first", "exactlyOne", "all"]).default("all"),
        ...page,
      })
      .strict(),
    output: docxAgentResultSchemas.read,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2", "guarded-ooxml"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_apply: {
    description: "Apply one action or workflow through a private candidate and durable receipt. Discover exact action arguments with docx_inspect.actions before first use; do not guess them. Reuse an opaque selectorTarget from query/read, or provide an inline select. Table workflows require a table selectorTarget or inline table select; use update_table_by_key for a row key and column header, including unmerged value cells in tables with unrelated merges. docx_inspect.workflowLimits states the supported shapes; do not guess cell coordinates after a refusal. Failed or partial engine receipts are never accepted.",
    input: z
      .object({
        view: documentView,
        operationId: id.optional(),
        mode: z.enum(["preview", "apply"]).default("apply"),
        action: z
          .enum(DOCX_AGENT_ACTIONS)
          .describe(
            "One bounded high-level action. docx_inspect returns exact action help and the current supported catalogue.",
          )
          .optional(),
        arguments: argumentsObject.optional(),
        selectorTarget: z
          .object({ id })
          .strict()
          .optional()
          .describe(
            "Opaque target returned by docx_query or a domain read. Query/read again after every mutation.",
          ),
        selectorTargets: z.array(selectorTarget).min(1).max(100).optional(),
        select: querySelect.optional().describe(
          "Inline selector resolved with require=exactlyOne. Prefer opaque targets when already available.",
        ),
        story: storySelector.optional(),
        steps: z.array(actionStep).min(1).max(25).optional(),
        workflow: compositeWorkflow.optional(),
        expect: postconditions.optional(),
        readback: z.object({
          maxItems: z.number().int().min(1).max(20).default(5),
          maxChars: z.number().int().min(100).max(16000).default(2000),
        }).strict().optional().describe("Optional bounded body text-block/table-cell changes with before/after text and fresh whole-block/table targets. maxChars is the total text budget. Preview has no reusable targets. Formatting, images and other stories/domains require explicit reads."),
        author: z.string().min(1).max(100).default("OfficeMCP"),
      })
      .strict()
      .superRefine((value, context) => {
        if (
          [Boolean(value.action), Boolean(value.steps), Boolean(value.workflow)].filter(Boolean)
            .length !== 1
        )
          context.addIssue({
            code: "custom",
            message: "Provide exactly one action, steps or workflow.",
            path: ["action"],
          });
        if (value.mode === "apply" && !value.operationId)
          context.addIssue({
            code: "custom",
            message: "operationId is required when mode is apply.",
            path: ["operationId"],
          });
      }),
    output: docxAgentResultSchemas.apply,
    risk: "mutate",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2", "guarded-ooxml"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_validate: {
    description: "Validate an unpublished candidate without publishing it. Returns SuperDoc reopen and independent package evidence.",
    input: z.object({ view: documentView }).strict(),
    output: docxAgentResultSchemas.validation,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2", "guarded-ooxml"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_compare: {
    description: "Compare the current source/candidate view with its opened source baseline using SuperDoc snapshots. Returns covered components and a structured change summary without modifying either document.",
    input: z.object({ view: documentView }).strict(),
    output: docxAgentResultSchemas.compare,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_image_get: {
    description: "Extract the exact PNG/JPEG bytes for an image target returned by docx_read(domain=images). Returns an owned asset and native MCP image content without executing embedded data.",
    input: z
      .object({ view: documentView, target: z.object({ id }).strict() })
      .strict(),
    output: docxAgentResultSchemas.media,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON + artifact + native MCP image","wireTypes":["structuredContent: application/json","content: matching JSON text","content: native PNG/JPEG image"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_export: {
    description: "Publish the exact DOCX view as a PDF through configured LibreOffice. The result identifies the conversion engine and does not claim Microsoft Word layout fidelity. Repeating operationId recovers the same publication result.",
    input: z
      .object({
        view: documentView,
        operationId: id,
        format: z.literal("pdf"),
        destination: z
          .object({ rootId: id, relativePath: z.string().min(1).max(4096) })
          .strict(),
      })
      .strict(),
    output: docxAgentResultSchemas.export,
    risk: "publish",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "taskEligible",
    engines: ["libreoffice","mupdf"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_save: {
    description: "Validate and save a candidate returned by docx_apply or docx_create as a new DOCX without overwrite. An open source view alone is not a candidate: for an unchanged copy, first call docx_create with action=from_template and path. If a Save response is lost, use docx_operation_get with the original operationId; repeating identical Save arguments also recovers its receipt. The original is never changed.",
    input: z.object({
      view: documentView,
      operationId: id,
      destination: z
        .object({ rootId: id, relativePath: z.string().min(1).max(4096) })
        .strict(),
    }),
    output: docxAgentResultSchemas.export,
    risk: "publish",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "synchronous",
    engines: ["superdoc-v2", "guarded-ooxml"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_close: {
    description: "Close a DOCX handle and optionally discard unpublished candidates. Closing never deletes the source or a published output.",
    input: z
      .object({
        view: documentView,
        disposition: z.enum(["retain", "discard"]).default("retain"),
      })
      .strict(),
    output: docxAgentResultSchemas.generic,
    risk: "mutate",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "synchronous",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
  docx_operation_get: {
    description: "Recover a durable DOCX mutation or publication receipt by operationId. Use the same operation ID after an uncertain response; never invent a new ID for the same write.",
    input: z.object({ operationId: id }).strict(),
    output: docxAgentResultSchemas.apply,
    risk: "read",
    handler: "docx.command",
    execute: (service, name, args, signal) => service.commands.docxCommand(name, args, signal),
    execution: "synchronous",
    engines: ["superdoc-v2"],
    presets: ["docx-agent"],
    discovery: {"resultType":"JSON","wireTypes":["structuredContent: application/json","content: matching JSON text"],"status":"Implemented · agent facade","preset":"docx-agent","contractVersion":"docx/6"},
  },
};

export const docxAgentTools = Object.freeze(docxAgentDefinitions);
