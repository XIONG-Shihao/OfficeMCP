import { z } from "zod";

const id = z.string().min(1).max(1000);
const text = z.string().max(32000);
const nonempty = text.min(1);
const index = z.number().int().min(0);
const ordinal = z.number().int().min(1);
const changeMode = z.enum(["direct", "tracked"]).default("direct");
const color = z.string().min(1).max(32);
const placement = z
  .object({
    at: z.enum(["document_start", "document_end", "before", "after"]),
  })
  .strict();
const empty = z.object({}).strict();
const tableCell = z
  .object({ row: index, column: index, expectedText: text, text })
  .strict();
const border = z
  .object({
    edge: z.enum([
      "top",
      "bottom",
      "left",
      "right",
      "insideH",
      "insideV",
      "between",
      "bar",
    ]),
    lineStyle: z.string().min(1).max(50).default("single"),
    lineWeightPt: z.number().min(0.25).max(24).default(0.5),
    color,
  })
  .strict();

const registry = {};
export const sectionBlocks = z.array(z.discriminatedUnion('kind', [
  z.object({ kind: z.literal('paragraph'), text }).strict(),
  z.object({ kind: z.literal('heading'), text: nonempty, level: z.number().int().min(1).max(6) }).strict(),
  z.object({ kind: z.literal('list'), items: z.array(nonempty).min(1).max(100), listKind: z.enum(['ordered', 'bullet']).default('bullet') }).strict(),
  z.object({ kind: z.literal('table'), rows: z.array(z.array(text).min(1).max(20)).min(1).max(100) }).strict(),
  z.object({ kind: z.literal('retain'), index: z.number().int().min(0).max(99) }).strict(),
])).min(1).max(100);
function add(
  name,
  description,
  schema,
  {
    target = "none",
    cardinality = "one",
    handler = "core",
    atomicity = "superdoc-action",
  } = {},
) {
  registry[name] = {
    name,
    description,
    schema,
    target,
    cardinality,
    handler,
    atomicity,
    risk: "mutate",
    execution: "taskEligible",
    supportedEngines: ["superdoc-v2"],
  };
}

add(
  "insert_paragraphs",
  "Insert one or more paragraphs, optionally making the first a heading.",
  z
    .object({
      texts: z.array(nonempty).min(1).max(100).optional(),
      text: nonempty.optional(),
      headingLevel: z.number().int().min(1).max(6).optional(),
      placement: placement.optional(),
      changeMode,
    })
    .strict()
    .refine((a) => Boolean(a.text) !== Boolean(a.texts), {
      message: "Provide exactly one of text or texts.",
    }),
  { target: "block-optional" },
);
add(
  "insert_heading",
  "Insert a heading at an explicit or inferred placement.",
  z
    .object({
      text: nonempty,
      level: z.number().int().min(1).max(6),
      placement: placement.optional(),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add(
  "replace_text",
  "Replace bounded literal text edits, optionally within one selected block.",
  z
    .object({
      edits: z
        .array(z.object({ find: nonempty, replace: text }).strict())
        .min(1)
        .max(100),
      caseSensitive: z.boolean().default(false),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add(
  "delete_text",
  "Delete bounded literal strings, optionally within one selected block.",
  z
    .object({
      finds: z.array(nonempty).min(1).max(100),
      caseSensitive: z.boolean().default(false),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add("delete_blocks", "Delete one or more selected blocks.", z.object({ changeMode }).strict(), {
  target: "block",
  cardinality: "many",
});
add(
  "append_list",
  "Create a new ordered or bullet list.",
  z
    .object({
      items: z.array(nonempty).min(1).max(100),
      kind: z.enum(["ordered", "bullet"]).default("bullet"),
      headingText: nonempty.optional(),
      headingLevel: z.number().int().min(1).max(6).optional(),
      placement: placement.optional(),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add(
  "create_table",
  "Create a bounded table with optional initial text.",
  z
    .object({
      rows: z.number().int().min(1).max(100),
      columns: z.number().int().min(1).max(20),
      cellTexts: z.array(z.array(text).max(20)).max(100).optional(),
      placement: placement.optional(),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add("rewrite_block", "Rewrite exactly one selected block.", z.object({ text: nonempty, changeMode }).strict(), {
  target: "block",
});
for (const name of ["accept_tracked_changes", "reject_tracked_changes"])
  add(
    name,
    `${name.startsWith("accept") ? "Accept" : "Reject"} tracked changes by optional author and type. Complex formatted or table-bearing ranges may be refused by the engine without changing the document.`,
    z
      .object({
        author: z.string().min(1).max(100).optional(),
        changeType: z.enum(["insert", "delete", "replacement", "format"]).optional(),
      })
      .strict(),
    {},
  );
add(
  "normalize_body_font_size",
  "Set the body font size.",
  z.object({ fontSize: z.number().min(1).max(200), changeMode }).strict(),
);
add(
  "set_font_family",
  "Set a font family for a block, literal text, or the body.",
  z
    .object({
      fontFamily: z.string().min(1).max(100),
      targetText: nonempty.optional(),
      targetTexts: z.array(nonempty).min(1).max(100).optional(),
      caseSensitive: z.boolean().default(false),
      changeMode,
    })
    .strict(),
  { target: "block-optional" },
);
add(
  "apply_letter_spacing",
  "Set letter spacing on one selected block.",
  z.object({ letterSpacing: z.number().min(-20).max(100), changeMode }).strict(),
  { target: "block" },
);
add("insert_toc", "Insert a table of contents field.", z.object({ title: text.optional(), placement: placement.optional(), changeMode }).strict(), {
  target: "block-optional",
});

for (const [name, shape] of [
  ["insert_table_row", { tableOrdinal: ordinal.optional(), rowIndex: index.optional(), position: z.enum(["before", "after", "above", "below"]).optional(), cellTexts: z.array(text).max(20).optional(), changeMode, dryRun: z.boolean().default(false) }],
  ["insert_table_column", { tableOrdinal: ordinal.optional(), columnIndex: index.optional(), position: z.enum(["left", "right"]).optional(), headerText: text.optional(), changeMode }],
  ["delete_table_row", { tableOrdinal: ordinal.optional(), rowIndex: index, changeMode }],
  ["delete_table_column", { tableOrdinal: ordinal.optional(), columnIndex: index, changeMode }],
  ["split_table", { tableOrdinal: ordinal.optional(), rowIndex: z.number().int().min(1), separatorText: text.optional(), changeMode }],
])
  add(name, `${name.replaceAll("_", " ")}.`, z.object(shape).strict(), {
    target: "table-optional",
  });

add(
  "convert_list",
  "Convert an existing list or paragraph range to ordered or bullet form.",
  z.object({ kind: z.enum(["ordered", "bullet"]), listOrdinal: ordinal.optional(), anchorText: nonempty.optional(), fromMarker: nonempty.optional(), toMarker: nonempty.optional(), fromText: nonempty.optional(), toText: nonempty.optional(), changeMode }).strict(),
);
add("split_list", "Split a list at the identified item.", z.object({ anchorText: nonempty, restartNumbering: z.boolean().default(true) }).strict());
add("attach_numbering", "Attach numbering matching an existing marker.", z.object({ nodeId: id.optional(), anchorText: nonempty.optional(), likeMarker: nonempty }).strict(), {
  target: "block-optional",
});
add(
  "add_list_items",
  "Add items to an existing list while preserving its numbering.",
  z.object({ anchorText: nonempty.optional(), listOrdinal: ordinal.optional(), entries: z.array(z.object({ text: nonempty, level: z.number().int().min(0).max(8).optional() }).strict()).min(1).max(100).optional(), items: z.array(nonempty).min(1).max(100).optional(), changeMode }).strict(),
);
add(
  "format_text",
  "Apply one inline-format patch to literal text or one selected block.",
  z.object({ targetText: nonempty.optional(), targetTexts: z.array(nonempty).min(1).max(100).optional(), caseSensitive: z.boolean().default(false), bold: z.boolean().optional(), italic: z.boolean().optional(), underline: z.boolean().optional(), strike: z.boolean().optional(), highlight: color.optional(), color: color.optional(), fontSize: z.number().min(1).max(200).optional(), changeMode }).strict(),
  { target: "block-optional" },
);
add(
  "apply_style",
  "Apply a style ID, heading level, or another block's effective style.",
  z.object({ styleId: id.optional(), headingLevel: z.number().int().min(1).max(6).optional(), likeText: nonempty.optional() }).strict().refine((a) => [a.styleId, a.headingLevel, a.likeText].filter(Boolean).length === 1, {
    message: "Provide exactly one styleId, headingLevel or likeText.",
  }),
  { target: "block" },
);
add("format_paragraph", "Set alignment on one selected paragraph.", z.object({ alignment: z.enum(["left", "center", "right", "justify"]), changeMode }).strict(), { target: "block" });
add("style_table", "Apply the qualified professional table style.", z.object({ tableOrdinal: ordinal.optional(), accentColor: color.optional() }).strict(), { target: "table-optional" });
add("move_table", "Move a whole table to a bounded placement.", z.object({ tableOrdinal: ordinal.optional(), placement }).strict(), { target: "table-optional" });
add("delete_table", "Delete a whole table.", z.object({ tableOrdinal: ordinal.optional(), changeMode }).strict(), { target: "table-optional" });
add(
  "set_paragraph_spacing",
  "Set line and before/after spacing on one paragraph.",
  z.object({ lineSpacing: z.number().min(0.5).max(10).optional(), spaceBefore: z.number().min(0).max(2000).optional(), spaceAfter: z.number().min(0).max(2000).optional() }).strict(),
  { target: "block" },
);
add("insert_page_break", "Start a selected block on a new page.", empty, { target: "block" });
add("add_hyperlink", "Wrap exact existing text in an HTTP(S) or mailto hyperlink.", z.object({ text: nonempty, url: nonempty, tooltip: text.optional() }).strict());

const extended = (name, description, schema, target = "none") =>
  add(name, description, schema, { target, handler: "extended", atomicity: "candidate-transaction" });
extended("write_table_cells", "Write 1-100 guarded plain table cells.", z.object({ cells: z.array(tableCell).min(1).max(100) }).strict(), "table");
extended("merge_table_cells", "Merge a rectangular table-cell range.", z.object({ start: z.object({ row: index, column: index }).strict(), end: z.object({ row: index, column: index }).strict() }).strict(), "table");
extended("unmerge_table_cells", "Unmerge the table cell at a coordinate.", z.object({ row: index, column: index }).strict(), "table");
extended(
  "insert_image",
  "Insert an allowed PNG/JPEG asset.",
  z.object({ assetId: id.optional(), path: z.string().min(1).max(4096).optional(), alt: text.default(""), width: z.number().min(1).max(2000).optional(), height: z.number().min(1).max(2000).optional(), placement: z.enum(["start", "end", "before", "after"]).default("end") }).strict().refine((a) => Boolean(a.assetId) !== Boolean(a.path), { message: "Provide exactly one assetId or path." }),
  "block-optional",
);
extended(
  "replace_image",
  "Replace selected image bytes.",
  z.object({ assetId: id.optional(), path: z.string().min(1).max(4096).optional(), resetSize: z.boolean().default(false) }).strict().refine((a) => Boolean(a.assetId) !== Boolean(a.path), { message: "Provide exactly one assetId or path." }),
  "image",
);
extended("delete_image", "Delete one selected image.", empty, "image");
extended("resize_image", "Set selected image dimensions.", z.object({ width: z.number().min(1).max(10000), height: z.number().min(1).max(10000), unit: z.enum(["pt", "px", "twip"]).default("pt") }).strict(), "image");
extended(
  "crop_image",
  "Set selected image crop percentages.",
  z
    .object({
      leftPercent: z.number().min(0).max(100).optional(),
      topPercent: z.number().min(0).max(100).optional(),
      rightPercent: z.number().min(0).max(100).optional(),
      bottomPercent: z.number().min(0).max(100).optional(),
    })
    .strict()
    .refine((value) => Object.values(value).some((item) => item !== undefined), {
      message: "Provide at least one crop edge.",
    }),
  "image",
);
extended(
  "position_image",
  "Set selected floating-image anchors, alignment and point offsets.",
  z
    .object({
      horizontalRelativeTo: z.enum(["page", "margin", "column", "character"]),
      verticalRelativeTo: z.enum(["page", "margin", "paragraph", "line"]),
      horizontalAlign: z.enum(["left", "center", "right", "inside", "outside"]).optional(),
      verticalAlign: z.enum(["top", "center", "bottom", "inside", "outside"]).optional(),
      offsetPt: z
        .object({ x: z.number().min(-10000).max(10000), y: z.number().min(-10000).max(10000) })
        .strict()
        .optional(),
    })
    .strict(),
  "image",
);
extended(
  "set_image_wrap",
  "Set selected image wrapping with typed distances in points.",
  z
    .object({
      type: z.enum(["inline", "square", "tight", "through", "topAndBottom", "behindText", "inFrontOfText"]),
      side: z.enum(["both", "left", "right", "largest"]).optional(),
      distancesPt: z
        .object({
          top: z.number().min(0).max(1000).optional(),
          bottom: z.number().min(0).max(1000).optional(),
          left: z.number().min(0).max(1000).optional(),
          right: z.number().min(0).max(1000).optional(),
        })
        .strict()
        .optional(),
    })
    .strict(),
  "image",
);
extended("set_image_alt_text", "Set selected image alternative text.", z.object({ description: text }).strict(), "image");
extended("rotate_image", "Rotate one selected image by degrees.", z.object({ angle: z.number().min(-360).max(360) }).strict(), "image");
for (const decision of ["accept", "reject"])
  extended(
    `${decision}_tracked_change`,
    `${decision === "accept" ? "Accept" : "Reject"} one explicitly selected text tracked change. Formatting-only and structurally invalid ranges are refused without changing the document.`,
    empty,
    "change",
  );
extended("insert_section_break", "Insert a section break at a bounded placement.", z.object({ placement: z.enum(["start", "end", "before", "after"]), breakType: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).default("nextPage") }).strict(), "block-optional");
extended(
  "set_section_layout",
  "Set page size, orientation, margins or break type for one selected section using point units.",
  z
    .object({
      widthPt: z.number().min(72).max(10000).optional(),
      heightPt: z.number().min(72).max(10000).optional(),
      orientation: z.enum(["portrait", "landscape"]).optional(),
      marginsPt: z
        .object({
          top: z.number().min(0).max(2000).optional(),
          right: z.number().min(0).max(2000).optional(),
          bottom: z.number().min(0).max(2000).optional(),
          left: z.number().min(0).max(2000).optional(),
        })
        .strict()
        .optional(),
      breakType: z.enum(["nextPage", "continuous", "evenPage", "oddPage"]).optional(),
    })
    .strict()
    .refine((value) => Object.values(value).some((item) => item !== undefined), {
      message: "Provide page size, orientation, margins or breakType.",
    }),
  "section",
);
extended("set_content_control_text", "Set one guarded plain-text content control.", z.object({ expectedText: text, text }).strict(), "control");
extended("set_checkbox_control", "Set one checkbox content control.", z.object({ checked: z.boolean() }).strict(), "control");
extended("set_date_control", "Set one date content control.", z.object({ value: z.string().min(1).max(100) }).strict(), "control");
extended("set_choice_control", "Set one dropdown or combo-box selection.", z.object({ value: text }).strict(), "control");
extended(
  "edit_repeating_section",
  "Clone, insert or delete one item in a repeating-section content control.",
  z
    .object({
      action: z.enum(["clone", "insertBefore", "insertAfter", "delete"]),
      index: z.number().int().min(0).max(10000),
      changeMode,
    })
    .strict(),
  "control",
);
extended("update_toc", "Refresh all table-of-contents entries.", empty, "none");
extended(
  "format_table_cell",
  "Apply cell width, alignment, wrapping, padding, shading and one or more borders in one operation.",
  z
    .object({
      preferredWidthPt: z.number().min(0).max(2000).optional(),
      preferredWidthType: z.enum(["auto", "dxa", "nil", "pct"]).optional(),
      verticalAlign: z.enum(["top", "center", "bottom"]).optional(),
      wrapText: z.boolean().optional(),
      fitText: z.boolean().optional(),
      padding: z
        .object({
          topPt: z.number().min(0).max(200).optional(),
          rightPt: z.number().min(0).max(200).optional(),
          bottomPt: z.number().min(0).max(200).optional(),
          leftPt: z.number().min(0).max(200).optional(),
        })
        .strict()
        .optional(),
      shading: color.optional(),
      borders: z.array(border).min(1).max(8).optional(),
      changeMode,
    })
    .strict()
    .refine(
      (a) =>
        [
          a.preferredWidthPt,
          a.preferredWidthType,
          a.verticalAlign,
          a.wrapText,
          a.fitText,
          a.padding,
          a.shading,
          a.borders,
        ].some((value) => value !== undefined),
      { message: "Provide at least one cell formatting property." },
    ),
  "tableCell",
);
extended(
  "format_table_row",
  "Set row height, page-break behavior and repeating-header behavior.",
  z
    .object({
      heightPt: z.number().min(0).max(2000).optional(),
      heightRule: z.enum(["auto", "atLeast", "exact"]).optional(),
      allowBreakAcrossPages: z.boolean().optional(),
      repeatHeader: z.boolean().optional(),
      changeMode,
    })
    .strict()
    .refine(
      (a) =>
        [a.heightPt, a.allowBreakAcrossPages, a.repeatHeader].some(
          (value) => value !== undefined,
        ),
      { message: "Provide row height or a row option." },
    ),
  "tableRow",
);
extended(
  "format_paragraph_advanced",
  "Apply indentation, a tab stop, borders and shading to one paragraph in one candidate transaction.",
  z
    .object({
      indentation: z
        .object({
          left: z.number().min(-2000).max(2000).optional(),
          right: z.number().min(-2000).max(2000).optional(),
          firstLine: z.number().min(-2000).max(2000).optional(),
          hanging: z.number().min(0).max(2000).optional(),
        })
        .strict()
        .optional(),
      tabStop: z
        .object({
          position: z.number().min(0).max(5000),
          alignment: z.enum(["left", "center", "right", "decimal", "bar"]),
          leader: z.enum(["none", "dot", "hyphen", "underscore", "heavy", "middleDot"]).optional(),
        })
        .strict()
        .optional(),
      borders: z.array(border).min(1).max(8).optional(),
      shading: z
        .object({ fill: color, color: color.optional(), pattern: z.string().min(1).max(50).optional() })
        .strict()
        .optional(),
      changeMode,
    })
    .strict()
    .refine((a) => Boolean(a.indentation || a.tabStop || a.borders || a.shading), {
      message: "Provide indentation, tabStop, borders or shading.",
    }),
  "block",
);
extended(
  "materialize_header_footer",
  "Give one section its own header/footer part, copying inherited content when present.",
  z
    .object({
      sectionId: id,
      kind: z.enum(["header", "footer"]),
      variant: z.enum(["default", "first", "even"]).default("default"),
      changeMode,
    })
    .strict(),
  "none",
);
extended(
  "replace_header_footer_text",
  "Materialize one inherited header/footer when needed, then replace exactly one literal range in that section slot.",
  z
    .object({
      sectionId: id,
      kind: z.enum(["header", "footer"]),
      variant: z.enum(["default", "first", "even"]).default("default"),
      find: nonempty,
      replace: text,
      materializeInherited: z.boolean().default(true),
      changeMode,
    })
    .strict(),
  "none",
);
extended(
  "patch_hyperlink",
  "Update the destination or metadata of one hyperlink.",
  z
    .object({
      href: text.optional(),
      anchor: text.optional(),
      docLocation: text.optional(),
      tooltip: text.optional(),
      target: text.optional(),
      rel: text.optional(),
      changeMode,
    })
    .strict()
    .refine(
      (a) => [a.href, a.anchor, a.docLocation, a.tooltip, a.target, a.rel].some((value) => value !== undefined),
      { message: "Provide at least one hyperlink patch property." },
    ),
  "hyperlink",
);
extended("remove_hyperlink", "Remove one hyperlink while preserving its display text.", z.object({ changeMode }).strict(), "hyperlink");
extended("rename_bookmark", "Rename one bookmark.", z.object({ newName: nonempty, changeMode }).strict(), "bookmark");
extended("remove_bookmark", "Remove one bookmark without deleting its text; refuse incoming references.", z.object({ incomingReferences: z.literal("refuse"), changeMode }).strict(), "bookmark");
extended("update_footnote", "Replace the plain text of one footnote.", z.object({ content: text, changeMode }).strict(), "footnote");
extended("remove_footnote", "Remove one footnote and its reference explicitly.", z.object({ referencePolicy: z.literal("remove_reference"), changeMode }).strict(), "footnote");
extended("remove_field", "Remove one raw field with an explicit cached-result policy.", z.object({ resultPolicy: z.enum(["remove", "keep_text"]), changeMode }).strict(), "field");
extended(
  "insert_caption",
  "Insert a numbered caption adjacent to one selected block.",
  z
    .object({
      position: z.enum(["above", "below"]).default("below"),
      label: nonempty.default("Figure"),
      text: text.optional(),
  })
    .strict(),
  "block",
);
extended(
  "update_caption",
  "Update the text of one selected caption.",
  z.object({ text, changeMode }).strict(),
  "caption",
);
extended(
  "remove_caption",
  "Remove one selected caption.",
  z.object({ changeMode }).strict(),
  "caption",
);
const crossReferenceTarget = z.discriminatedUnion("kind", [
  z.object({ kind: z.literal("bookmark"), name: nonempty }).strict(),
  z.object({ kind: z.literal("heading"), nodeId: id }).strict(),
  z.object({ kind: z.literal("note"), noteId: id }).strict(),
  z.object({ kind: z.literal("caption"), nodeId: id }).strict(),
  z.object({ kind: z.literal("numberedItem"), nodeId: id }).strict(),
]);
extended(
  "insert_cross_reference",
  "Insert a cross-reference at the selected text range.",
  z
    .object({
      reference: crossReferenceTarget,
      display: nonempty,
      changeMode,
    })
    .strict(),
  "block",
);
extended(
  "remove_cross_reference",
  "Remove one selected cross-reference with an explicit display-text policy.",
  z.object({ displayTextPolicy: z.enum(["remove", "keep_text"]), changeMode }).strict(),
  "crossRef",
);
extended(
  "update_endnote",
  "Replace the plain text of one endnote.",
  z.object({ content: text, changeMode }).strict(),
  "endnote",
);
extended(
  "remove_endnote",
  "Remove one endnote and its reference explicitly.",
  z.object({ referencePolicy: z.literal("remove_reference"), changeMode }).strict(),
  "endnote",
);

for (const [name, description, schema] of [
  ["replace_textbox_text", "Replace one literal range inside an existing textbox while preserving its surrounding OOXML and formatting. Guarded OOXML adapter; protected review/field content is refused.", z.object({ expectedText: text, find: nonempty, replace: text }).strict()],
  ["replace_equation_text", "Replace one literal symbol/text range inside existing OfficeMath without rebuilding its equation structure. Guarded OOXML adapter; does not evaluate mathematics.", z.object({ expectedText: text, find: nonempty, replace: text }).strict()],
  ["replace_table_cell_text", "Replace one literal range in an explicit merged owner or nested cell. Own cell paragraphs are joined with newline for expectedText. SDK range editing is tried first; guarded OOXML handles unexposed shapes.", z.object({ expectedText: text, find: text, replace: text, paragraphIndex: index.optional() }).strict()],
  ["set_chart_title", "Update an existing direct rich-text chart title. Computed or missing titles are refused; chart data is unchanged.", z.object({ expectedText: nonempty, text: nonempty }).strict()],
  ["set_chart_data", "Update one unique chart series and its embedded literal XLSX data atomically. Existing point count only; local simple A1 references or literal data. Formula, shared, external, pivot and irregular data are refused.", z.object({ seriesName: nonempty, expectedValues: z.array(z.number().finite()).min(1).max(1000), values: z.array(z.number().finite()).min(1).max(1000), expectedCategories: z.array(text).min(1).max(1000).optional(), categories: z.array(text).min(1).max(1000).optional() }).strict().refine(a => (a.categories === undefined) === (a.expectedCategories === undefined), { message: "Provide both expectedCategories and categories." })],
]) {
  add(name, description, schema, { target: "object", handler: "opc", atomicity: "isolated-validated-OOXML-transaction" });
  registry[name].supportedEngines = ["guarded-ooxml"];
}

registry.replace_table_cell_text.supportedEngines = ['superdoc-v2', 'guarded-ooxml'];
add('replace_section_content', 'Replace a rich section through guarded OOXML. Explicit retain indices preserve objects/review blocks exactly; comments and references cannot be implicitly deleted.', z.object({ headingPath: z.array(nonempty).min(1).max(20), subsections: z.enum(['preserve', 'replace']), expectedText: text, blocks: sectionBlocks }).strict(), { handler: 'opc', target: 'none', atomicity: 'isolated-validated-OOXML-transaction' });
registry.replace_section_content.supportedEngines = ['guarded-ooxml'];
const familyDefinitions = [
  ["objects", "Guarded edits of complex OOXML content outside native SDK mutation support.", ["replace_textbox_text", "replace_equation_text", "set_chart_title", "set_chart_data", "replace_section_content"]],
  ["text", "Insert, replace or remove paragraphs and literal text.", ["insert_paragraphs", "insert_heading", "replace_text", "delete_text", "delete_blocks", "rewrite_block"]],
  ["formatting", "Change character, paragraph and style formatting.", ["normalize_body_font_size", "set_font_family", "apply_letter_spacing", "format_text", "apply_style", "format_paragraph", "set_paragraph_spacing", "format_paragraph_advanced"]],
  ["lists", "Create or restructure lists and numbering.", ["append_list", "convert_list", "split_list", "attach_numbering", "add_list_items"]],
  ["tables", "Create, edit, arrange or format table structures and cells.", ["create_table", "insert_table_row", "insert_table_column", "delete_table_row", "delete_table_column", "split_table", "style_table", "move_table", "delete_table", "write_table_cells", "merge_table_cells", "unmerge_table_cells", "format_table_cell", "format_table_row", "replace_table_cell_text"]],
  ["images", "Insert, replace, arrange and describe images.", ["insert_image", "replace_image", "delete_image", "resize_image", "crop_image", "position_image", "set_image_wrap", "set_image_alt_text", "rotate_image"]],
  ["review", "Decide tracked changes. Existing comments are read-only through reads and queries.", ["accept_tracked_changes", "reject_tracked_changes", "accept_tracked_change", "reject_tracked_change"]],
  ["layout", "Set page, section and header/footer content or layout.", ["insert_page_break", "insert_section_break", "set_section_layout", "materialize_header_footer", "replace_header_footer_text"]],
  ["controls", "Fill typed fields and repeating content controls.", ["set_content_control_text", "set_checkbox_control", "set_date_control", "set_choice_control", "edit_repeating_section"]],
  ["links", "Maintain hyperlinks and bookmarks.", ["add_hyperlink", "patch_hyperlink", "remove_hyperlink", "rename_bookmark", "remove_bookmark"]],
  ["notes", "Update or remove footnotes and endnotes.", ["update_footnote", "remove_footnote", "update_endnote", "remove_endnote"]],
  ["references", "Maintain contents, captions, cross-references and fields.", ["insert_toc", "update_toc", "remove_field", "insert_caption", "update_caption", "remove_caption", "insert_cross_reference", "remove_cross_reference"]],
];

for (const [family, , actions] of familyDefinitions) {
  if (actions.length > 20) throw new Error(`DOCX action family ${family} exceeds the descriptor limit.`);
  for (const action of actions) {
    if (!registry[action] || registry[action].family)
      throw new Error(`Unknown or duplicate DOCX action family member: ${action}.`);
    registry[action].family = family;
  }
}
if (Object.values(registry).some((entry) => !entry.family))
  throw new Error("Every DOCX action must declare a discovery family.");

export const DOCX_ACTION_FAMILIES = Object.freeze(familyDefinitions.map(([id, description, actions]) =>
  Object.freeze({ id, description, actions: Object.freeze(actions) })));
export const DOCX_ACTION_FAMILY_IDS = Object.freeze(DOCX_ACTION_FAMILIES.map(({ id }) => id));
export const DOCX_ACTION_REGISTRY = Object.freeze(Object.fromEntries(
  Object.entries(registry).map(([name, entry]) => [name, Object.freeze(entry)]),
));
export const DOCX_AGENT_ACTIONS = Object.freeze(Object.keys(registry));
export const DOCX_ACTION_HELP = Object.fromEntries(
  Object.values(registry).map((entry) => [entry.name, entry.description]),
);

export function matchesDocxActionTarget(requirement, locator) {
  const kind = requirement.replace("-optional", "");
  if (kind === "none") return true;
  if (kind === "block")
    return (
      locator?.kind === "table" ||
      (locator?.kind === "agent_selector" &&
        [
          "paragraph", "heading", "listItem", "table", "tableRow",
          "tableCell", "tableOfContents", "sdt",
        ].includes(locator?.address?.nodeType))
    );
  if (kind === "object") return locator?.kind === "opc_object";
  if (kind === "comment") return locator?.kind === "agent_comment";
  if (["tableCell", "tableRow"].includes(kind))
    return locator?.kind === "agent_selector" && locator?.address?.nodeType === kind;
  return locator?.kind === kind;
}

export function parseDocxActionArguments(name, value) {
  const entry = registry[name];
  if (!entry) throw new Error(`Unknown DOCX action: ${name}.`);
  const parsed = entry.schema.safeParse(value ?? {});
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((issue) => `${issue.path.join(".") || "arguments"}: ${issue.message}`)
      .join("; ");
    throw new Error(`Invalid arguments for ${name}: ${issues}`);
  }
  return parsed.data;
}

/** Default discovery is compatible; a family request supplies at most 20 schemas. */
export function docxActionDiscovery({ actionFamily, actions } = {}) {
  const family = actionFamily && DOCX_ACTION_FAMILIES.find(({ id }) => id === actionFamily);
  if (actionFamily && !family) throw new Error(`Unknown DOCX action family: ${actionFamily}.`);
  const selected = actions ?? family?.actions ?? [];
  if (selected.length > 20) throw new Error("At most 20 DOCX action descriptors may be requested.");
  if (selected.some((action) => !registry[action])) throw new Error("Unknown DOCX action descriptor.");
  return {
    availableActions: family?.actions ?? DOCX_AGENT_ACTIONS,
    actionFamilies: DOCX_ACTION_FAMILIES.map(({ id, description, actions }) => ({ id, description, actionCount: actions.length })),
    ...(family ? { selectedActionFamily: family.id } : {}),
    actionDescriptors: Object.fromEntries(selected.map((name) => [name, docxActionDescriptor(name)])),
  };
}

export function docxActionDescriptor(name) {
  const entry = registry[name];
  if (!entry) return null;
  return {
    name,
    family: entry.family,
    description: entry.description,
    requiredTargetKind: entry.target,
    targetCardinality: entry.cardinality,
    risk: entry.risk,
    execution: entry.execution,
    supportedEngines: entry.supportedEngines,
    atomicity: entry.atomicity,
    argumentSchema: z.toJSONSchema(entry.schema, { target: "draft-2020-12" }),
  };
}
