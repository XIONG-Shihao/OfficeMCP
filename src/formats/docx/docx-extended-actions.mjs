import { replaceStoryText } from "./docx-text-actions.mjs";
import { decideTrackedChange } from "./docx-review-actions.mjs";
import fs from "node:fs";
import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";
import { list, success } from "./docx-sdk.mjs";
import { writeTableCells } from "./docx-table-actions.mjs";

const targeted = (invoke) => async ({ doc, args, locator }) => {
  return success(await invoke(doc, args, locator));
};

async function handle_write_table_cells({ doc, step, args, locator, placementOf }) {
  return writeTableCells(doc, locator, args.cells);
}

async function handle_insert_image({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(step.image?.file, "INVALID_ARGUMENT", "Provide an image assetId or path.");
  const bytes = fs.readFileSync(step.image.file);
  result = success(await doc.create.image({
        src: `data:${step.image.mediaType};base64,${bytes.toString("base64")}`,
        alt: args.alt ?? "",
        size: {
          ...(args.width !== undefined ? { width: args.width } : {}),
          ...(args.height !== undefined ? { height: args.height } : {}),
        },
        at: placementOf(args, locator),
      }));
  return result;
}

async function handle_replace_image({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(locator?.kind === "image" && step.image?.file, "INVALID_ARGUMENT", "Use an image selectorTarget and provide replacement image bytes.");
  const bytes = fs.readFileSync(step.image.file);
  result = success(await doc.images.replaceSource({
        imageId: locator.nodeId,
        src: `data:${step.image.mediaType};base64,${bytes.toString("base64")}`,
        resetSize: args.resetSize ?? false,
      }));
  return result;
}

async function handle_set_image_wrap({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(locator?.kind === "image", "INVALID_ARGUMENT", "Use an image selectorTarget.");
  success(await doc.images.setWrapType({ imageId: locator.nodeId, type: args.type }));
  if (args.side !== undefined)
        success(await doc.images.setWrapSide({ imageId: locator.nodeId, side: args.side }));
  if (args.distancesPt)
        success(await doc.images.setWrapDistances({
          imageId: locator.nodeId,
          distances: {
            distTop: args.distancesPt.top,
            distBottom: args.distancesPt.bottom,
            distLeft: args.distancesPt.left,
            distRight: args.distancesPt.right,
          },
        }));
  result = { success: true };
  return result;
}

async function handle_tracked_change_decision({ doc, step, locator }) {
  return decideTrackedChange(doc, locator, step.action.startsWith("accept") ? "accept" : "reject");
}

async function handle_insert_section_break({ doc, step, args, locator, placementOf }) {
  let result;
  result = success(await doc.create.sectionBreak({
        at: placementOf(args, locator),
        breakType: args.breakType ?? "nextPage",
      }));
  return result;
}

async function handle_set_section_layout({ doc, args, locator }) {
  requireThat(locator?.kind === "section", "INVALID_ARGUMENT", "Use one section target.");
  const target = locator.address;
  const operations = [];
  if (args.widthPt !== undefined || args.heightPt !== undefined || args.orientation)
    operations.push(
      success(
        await doc.sections.setPageSetup({
          target,
          ...(args.widthPt !== undefined ? { width: args.widthPt / 72 } : {}),
          ...(args.heightPt !== undefined ? { height: args.heightPt / 72 } : {}),
          ...(args.orientation ? { orientation: args.orientation } : {}),
        }),
      ),
    );
  if (args.marginsPt)
    operations.push(
      success(
        await doc.sections.setPageMargins({
          target,
          ...Object.fromEntries(
            Object.entries(args.marginsPt).map(([name, value]) => [name, value / 72]),
          ),
        }),
      ),
    );
  if (args.breakType)
    operations.push(
      success(await doc.sections.setBreakType({ target, breakType: args.breakType })),
    );
  return { operations: operations.length };
}

async function handle_set_content_control_text({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(locator?.kind === "control", "INVALID_ARGUMENT", "Use a content-control selectorTarget.");
  const control = await doc.contentControls.get({ target: locator.address });
  requireThat(control.text === args.expectedText, "TARGET_CONFLICT", "Content-control text changed.");
  result = success(await doc.contentControls.text.setValue({ target: locator.address, value: args.text }));
  return result;
}

async function handle_edit_repeating_section({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(locator?.kind === "control", "INVALID_ARGUMENT", "Use a repeating-section content-control selectorTarget.");
  const methods = {
        clone: "cloneItem",
        insertBefore: "insertItemBefore",
        insertAfter: "insertItemAfter",
        delete: "deleteItem",
      };
  result = success(
        await doc.contentControls.repeatingSection[methods[args.action]]({
          target: locator.address,
          index: args.index,
          changeMode: args.changeMode,
        }),
      );
  return result;
}

async function handle_update_toc({ doc, step, args, locator, placementOf }) {
  let result;
  const tablesOfContents = await list(doc, "toc.list");
  requireThat(tablesOfContents.length > 0, "NOT_FOUND", "No table of contents is available.");
  const updated = [];
  for (const toc of tablesOfContents) {
        const target = toc.target ?? toc.address ?? {
          kind: "block",
          nodeType: "tableOfContents",
          nodeId: toc.nodeId ?? toc.id,
        };
        updated.push(success(await doc.toc.update({ target })));
      }
  result = { updated: updated.length };
  return result;
}

async function handle_format_table_cell({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(
        locator?.kind === "agent_selector" && locator.address?.nodeType === "tableCell",
        "INVALID_ARGUMENT",
        "Use a table-cell selectorTarget.",
      );
  const target = locator.address;
  const operations = [];
  const properties = Object.fromEntries(
        ["preferredWidthPt", "preferredWidthType", "verticalAlign", "wrapText", "fitText"]
          .filter((key) => args[key] !== undefined)
          .map((key) => [key, args[key]]),
      );
  if (Object.keys(properties).length)
        operations.push(success(await doc.tables.setCellProperties({ target, ...properties, changeMode: args.changeMode })));
  if (args.padding)
        operations.push(success(await doc.tables.setCellPadding({ target, ...args.padding, changeMode: args.changeMode })));
  if (args.shading !== undefined)
        operations.push(success(await doc.tables.setShading({ target, color: args.shading, changeMode: args.changeMode })));
  for (const border of args.borders ?? [])
        operations.push(success(await doc.tables.setBorder({ target, ...border, changeMode: args.changeMode })));
  result = { operations: operations.length };
  return result;
}

async function handle_format_table_row({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(
        locator?.kind === "agent_selector" && locator.address?.nodeType === "tableRow",
        "INVALID_ARGUMENT",
        "Use a table-row selectorTarget.",
      );
  const target = locator.address;
  const operations = [];
  if (args.heightPt !== undefined)
        operations.push(success(await doc.tables.setRowHeight({
          target,
          heightPt: args.heightPt,
          rule: args.heightRule,
          changeMode: args.changeMode,
        })));
  if (args.allowBreakAcrossPages !== undefined || args.repeatHeader !== undefined)
        operations.push(success(await doc.tables.setRowOptions({
          target,
          allowBreakAcrossPages: args.allowBreakAcrossPages,
          repeatHeader: args.repeatHeader,
          changeMode: args.changeMode,
        })));
  result = { operations: operations.length };
  return result;
}

async function handle_format_paragraph_advanced({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(
        locator?.kind === "agent_selector" &&
          ["paragraph", "heading", "listItem"].includes(locator.address?.nodeType),
        "INVALID_ARGUMENT",
        "Use a paragraph, heading or list-item selectorTarget.",
      );
  const target = locator.address;
  const operations = [];
  if (args.indentation)
        operations.push(success(await doc.format.paragraph.setIndentation({ target, ...args.indentation, changeMode: args.changeMode })));
  if (args.tabStop)
        operations.push(success(await doc.format.paragraph.setTabStop({ target, ...args.tabStop, changeMode: args.changeMode })));
  for (const border of args.borders ?? [])
        operations.push(success(await doc.format.paragraph.setBorder({
          target,
          side: border.edge,
          style: border.lineStyle,
          size: Math.round(border.lineWeightPt * 8),
          color: border.color,
          changeMode: args.changeMode,
        })));
  if (args.shading)
        operations.push(success(await doc.format.paragraph.setShading({ target, ...args.shading, changeMode: args.changeMode })));
  result = { operations: operations.length };
  return result;
}

async function handle_materialize_header_footer({ doc, step, args, locator, placementOf }) {
  let result;
  const target = {
        kind: "headerFooterSlot",
        section: { kind: "section", sectionId: args.sectionId },
        headerFooterKind: args.kind,
        variant: args.variant,
      };
  const resolved = await doc.headerFooters.resolve({ target });
  if (resolved.status === "explicit") {
        result = { materialized: false, alreadyExplicit: true, refId: resolved.refId };
      } else {
        const created = success(await doc.headerFooters.parts.create({
          kind: args.kind,
          ...(resolved.status === "inherited" ? { sourceRefId: resolved.refId } : {}),
          changeMode: args.changeMode,
        }));
        success(await doc.headerFooters.refs.set({ target, refId: created.refId, changeMode: args.changeMode }));
        result = {
          materialized: true,
          copiedInheritedContent: resolved.status === "inherited",
          refId: created.refId,
        };
      }
  return result;
}

async function handle_replace_header_footer_text({ doc, step, args, locator, placementOf }) {
  let result;
  const target = {
        kind: "headerFooterSlot",
        section: { kind: "section", sectionId: args.sectionId },
        headerFooterKind: args.kind,
        variant: args.variant,
      };
  let resolved = await doc.headerFooters.resolve({ target });
  if (resolved.status === "inherited" && args.materializeInherited) {
        const created = success(await doc.headerFooters.parts.create({
          kind: args.kind,
          sourceRefId: resolved.refId,
          changeMode: args.changeMode,
        }));
        success(await doc.headerFooters.refs.set({
          target,
          refId: created.refId,
          changeMode: args.changeMode,
        }));
        resolved = { status: "explicit", refId: created.refId };
      }
  requireThat(
        resolved.status === "explicit" && resolved.refId,
        "UNSUPPORTED_FEATURE",
        "Header/footer slot has no explicit content. Enable materializeInherited or create the part first.",
      );
  const story = {
        kind: "story",
        storyType: "headerFooterPart",
        refId: resolved.refId,
      };
  result = await replaceStoryText(doc, { story, find: args.find, replace: args.replace, changeMode: args.changeMode });
  return result;
}

async function handle_remove_bookmark({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(locator?.kind === "bookmark", "INVALID_ARGUMENT", "Use a bookmark selectorTarget.");
  const references = await list(doc, "crossRefs.list");
  const name = locator.address?.name ?? locator.nodeId;
  const incoming = references.filter((reference) =>
    JSON.stringify(reference).includes(String(name)),
  );
  requireThat(
    incoming.length === 0,
    "TARGET_CONFLICT",
    `Bookmark has ${incoming.length} incoming cross-reference(s); remove or retarget them first.`,
  );
  result = success(await doc.bookmarks.remove({ target: locator.address, changeMode: args.changeMode }));
  return result;
}

async function handle_insert_caption({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(
        ["agent_selector", "table"].includes(locator?.kind) && locator.address?.nodeId,
        "INVALID_ARGUMENT",
        "Use a block selectorTarget for caption insertion.",
      );
  result = success(
        await doc.captions.insert({
          adjacentTo: locator.address,
          position: args.position,
          label: args.label,
          text: args.text,
        }),
      );
  return result;
}

async function handle_insert_cross_reference({ doc, step, args, locator, placementOf }) {
  let result;
  requireThat(
        locator?.kind === "agent_selector" && locator.textTarget,
        "INVALID_ARGUMENT",
        "Use a text query selectorTarget as the cross-reference insertion range.",
      );
  result = success(await doc.crossRefs.insert({
        at: locator.textTarget,
        target: args.reference,
        display: args.display,
        changeMode: args.changeMode,
      }));
  return result;
}

const EXTENDED_ACTION_HANDLERS = new Map([
  ["write_table_cells", handle_write_table_cells],
  ["merge_table_cells", targeted(
    (doc, args, locator) => doc.tables.mergeCells({
      target: locator.address,
      start: { rowIndex: args.start?.row, columnIndex: args.start?.column },
      end: { rowIndex: args.end?.row, columnIndex: args.end?.column },
    }))],
  ["unmerge_table_cells", targeted(
    (doc, args, locator) => doc.tables.unmergeCells({
      target: locator.address, rowIndex: args.row, columnIndex: args.column,
    }))],
  ["insert_image", handle_insert_image],
  ["replace_image", handle_replace_image],
  ["delete_image", targeted(
    (doc, _args, locator) => doc.images.delete({ imageId: locator.nodeId }))],
  ["resize_image", targeted(
    (doc, args, locator) => doc.images.setSize({
      imageId: locator.nodeId,
      size: { width: args.width, height: args.height, unit: args.unit ?? "pt" },
    }))],
  ["crop_image", targeted(
    (doc, args, locator) => doc.images.crop({
      imageId: locator.nodeId,
      crop: {
        left: args.leftPercent, top: args.topPercent,
        right: args.rightPercent, bottom: args.bottomPercent,
      },
    }))],
  ["position_image", targeted(
    (doc, args, locator) => doc.images.setPosition({
      imageId: locator.nodeId,
      position: {
        hRelativeFrom: args.horizontalRelativeTo,
        vRelativeFrom: args.verticalRelativeTo,
        alignH: args.horizontalAlign,
        alignV: args.verticalAlign,
        marginOffset: args.offsetPt
          ? { horizontal: args.offsetPt.x, top: args.offsetPt.y }
          : undefined,
      },
    }))],
  ["set_image_wrap", handle_set_image_wrap],
  ["set_image_alt_text", targeted(
    (doc, args, locator) => doc.images.setAltText({
      imageId: locator.nodeId, description: args.description,
    }))],
  ["rotate_image", targeted(
    (doc, args, locator) => doc.images.rotate({ imageId: locator.nodeId, angle: args.angle }))],
  ["accept_tracked_change", handle_tracked_change_decision],
  ["reject_tracked_change", handle_tracked_change_decision],
  ["insert_section_break", handle_insert_section_break],
  ["set_section_layout", handle_set_section_layout],
  ["set_content_control_text", handle_set_content_control_text],
  ["set_checkbox_control", targeted(
    (doc, args, locator) => doc.contentControls.checkbox.setState({
      target: locator.address, checked: args.checked,
    }))],
  ["set_date_control", targeted(
    (doc, args, locator) => doc.contentControls.date.setValue({
      target: locator.address, value: args.value,
    }))],
  ["set_choice_control", targeted(
    (doc, args, locator) => doc.contentControls.choiceList.setSelected({
      target: locator.address, value: args.value,
    }))],
  ["edit_repeating_section", handle_edit_repeating_section],
  ["update_toc", handle_update_toc],
  ["format_table_cell", handle_format_table_cell],
  ["format_table_row", handle_format_table_row],
  ["format_paragraph_advanced", handle_format_paragraph_advanced],
  ["materialize_header_footer", handle_materialize_header_footer],
  ["replace_header_footer_text", handle_replace_header_footer_text],
  ["patch_hyperlink", targeted(
    (doc, { changeMode, ...patch }, locator) =>
      doc.hyperlinks.patch({ target: locator.address, patch, changeMode }))],
  ["remove_hyperlink", targeted(
    (doc, args, locator) => doc.hyperlinks.remove({
      target: locator.address, mode: "keepText", changeMode: args.changeMode,
    }))],
  ["rename_bookmark", targeted(
    (doc, args, locator) => doc.bookmarks.rename({
      target: locator.address, newName: args.newName, changeMode: args.changeMode,
    }))],
  ["remove_bookmark", handle_remove_bookmark],
  ["update_footnote", targeted(
    (doc, args, locator) => doc.footnotes.update({
      target: locator.address, patch: { content: args.content }, changeMode: args.changeMode,
    }))],
  ["remove_footnote", targeted(
    (doc, args, locator) => doc.footnotes.remove({
      target: locator.address, changeMode: args.changeMode,
    }))],
  ["remove_field", targeted(
    (doc, args, locator) => doc.fields.remove({
      target: locator.address,
      mode: args.resultPolicy === "keep_text" ? "keepText" : "raw",
      changeMode: args.changeMode,
    }))],
  ["insert_caption", handle_insert_caption],
  ["update_caption", targeted(
    (doc, args, locator) => doc.captions.update({
      target: locator.address, patch: { text: args.text }, changeMode: args.changeMode,
    }))],
  ["remove_caption", targeted(
    (doc, args, locator) => doc.captions.remove({
      target: locator.address, changeMode: args.changeMode,
    }))],
  ["insert_cross_reference", handle_insert_cross_reference],
  ["remove_cross_reference", targeted(
    (doc, args, locator) => doc.crossRefs.remove({
      target: locator.address,
      mode: args.displayTextPolicy === "keep_text" ? "keepText" : "raw",
      changeMode: args.changeMode,
    }))],
  ["update_endnote", targeted(
    (doc, args, locator) => doc.footnotes.update({
      target: locator.address, patch: { content: args.content }, changeMode: args.changeMode,
    }))],
  ["remove_endnote", targeted(
    (doc, args, locator) => doc.footnotes.remove({
      target: locator.address, changeMode: args.changeMode,
    }))],
]);

export async function applyExtendedAction(doc, step, placementOf) {
  const args = step.arguments;
  const locator = step.targetLocator;
  const handler = EXTENDED_ACTION_HANDLERS.get(step.action);
  if (!handler)
    throw new OfficeError(
      "UNSUPPORTED_FEATURE",
      `Unsupported extended DOCX action: ${step.action}.`,
    );
  const result = await handler({ doc, step, args, locator, placementOf });
  return {
    status: "ok",
    intent: step.action,
    selectedTargets: locator
      ? [locator.nodeId ?? locator.address?.nodeId]
      : [],
    executedOperations: [{ operationId: step.action, result }],
    verificationPassed: true,
    verification: [{ check: { kind: "sdk-success-receipt" }, passed: true }],
  };
}
