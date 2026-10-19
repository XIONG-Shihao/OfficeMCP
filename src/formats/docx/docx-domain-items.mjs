import { list } from "./docx-sdk.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";

/** Shared typed domain inventory and mutation locators for every DOCX route. */
export async function readDocxDomain(doc, domain) {
  const methods = {
    sections: "sections.list",
    images: "images.list",
    fields: "fields.list",
    footnotes: "footnotes.list",
    endnotes: "footnotes.list",
    hyperlinks: "hyperlinks.list",
    bookmarks: "bookmarks.list",
    content_controls: "contentControls.list",
    tracked_changes: "trackChanges.list",
    comments: "comments.list",
    lists: "lists.list",
    toc: "toc.list",
    captions: "captions.list",
    cross_references: "crossRefs.list",
  };
  requireThat(
    methods[domain],
    "INVALID_ARGUMENT",
    "Unsupported SDK read domain.",
  );
  let items = await list(doc, methods[domain], domain === "endnotes" ? { type: "endnote" } : {});
  if (!Array.isArray(items)) items = items.items ?? [];
  items = items.map((i, index) => ({
    ...i,
    kind: i.kind ?? domain,
    ...(domain === "sections"
      ? {
          nativeTarget: {
            kind: "section",
            address: i.address,
            nodeId: i.id,
            index,
          },
        }
      : domain === "images"
        ? {
          nativeTarget: {
            kind: "image",
            address: i.address,
            nodeId: i.sdImageId,
            partName: i.properties?.src,
            index,
          },
          }
        : domain === "content_controls"
          ? {
              nativeTarget: {
                kind: "control",
                address: i.target,
                nodeId: i.id,
                text: i.text,
              },
            }
          : domain === "tracked_changes"
            ? {
                nativeTarget: {
                  kind: "change",
                  address: i.address,
                  nodeId: i.id,
                  change: i,
                },
              }
            : domain === "comments"
              ? {
                  nativeTarget: {
                    kind: "agent_comment",
                    commentId: i.id ?? i.entityId,
                    nodeId: i.id ?? i.entityId,
                  },
                }
              : domain === "hyperlinks" && i.address
                ? {
                    nativeTarget: {
                      kind: "hyperlink",
                      address: i.address,
                      nodeId: `${i.address.anchor?.start?.blockId ?? "hyperlink"}:${i.address.anchor?.start?.offset ?? index}`,
                    },
                  }
                : domain === "bookmarks" && (i.target ?? i.address ?? i.name)
                  ? {
                      nativeTarget: {
                        kind: "bookmark",
                        address:
                          i.target ??
                          i.address ?? {
                            kind: "entity",
                            entityType: "bookmark",
                            name: i.name,
                          },
                        nodeId: i.name ?? i.id ?? `bookmark-${index}`,
                      },
                    }
                  : domain === "footnotes" && (i.target ?? i.address ?? i.noteId ?? i.id)
                    ? {
                        nativeTarget: {
                          kind: "footnote",
                          address:
                            i.target ??
                            i.address ?? {
                              kind: "entity",
                              entityType: "footnote",
                              noteId: String(i.noteId ?? i.id),
                            },
                          nodeId: String(i.noteId ?? i.id ?? index),
                        },
                      }
                    : domain === "endnotes" &&
                        (i.target ?? i.address ?? i.noteId ?? i.id)
                      ? {
                          nativeTarget: {
                            kind: "endnote",
                            address:
                              i.target ??
                              i.address ?? {
                                kind: "entity",
                                entityType: "footnote",
                                noteId: String(i.noteId ?? i.id),
                              },
                            nodeId: String(i.noteId ?? i.id ?? index),
                          },
                        }
                      : domain === "captions" && (i.target ?? i.address)
                        ? {
                            nativeTarget: {
                              kind: "caption",
                              address: i.target ?? i.address,
                              nodeId:
                                i.nodeId ??
                                i.target?.nodeId ??
                                i.address?.nodeId ??
                                `caption-${index}`,
                            },
                          }
                        : domain === "cross_references" &&
                            (i.target ?? i.address)
                          ? {
                              nativeTarget: {
                                kind: "crossRef",
                                address: i.target ?? i.address,
                                nodeId:
                                  i.nodeId ??
                                  i.target?.anchor?.start?.blockId ??
                                  `cross-ref-${index}`,
                              },
                            }
                    : domain === "fields" && (i.target ?? i.address)
                      ? {
                          nativeTarget: {
                            kind: "field",
                            address: i.target ?? i.address,
                            nodeId: i.id ?? i.fieldId ?? `field-${index}`,
                          },
                        }
            : {}),
  }));
  return items;
}
