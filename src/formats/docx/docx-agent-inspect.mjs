import { list } from "./docx-sdk.mjs";
import { inspectDocxPackage } from "./docx-package-inventory.mjs";
import { docxActionDiscovery } from "./docx-agent-actions.mjs";

const manualDomains = new Set([
  "footnotes",
  "endnotes",
  "toc",
  "captions",
  "crossReferences",
]);

const domainMap = {
  paragraphs: "linear",
  headings: "linear",
  tables: "tables",
  lists: "lists",
  images: "images",
  comments: "comments",
  trackedChanges: "tracked_changes",
  sections: "sections",
  fields: "fields",
  hyperlinks: "hyperlinks",
  bookmarks: "bookmarks",
  contentControls: "content_controls",
  headers: "headers_footers",
  footers: "headers_footers",
  footnotes: "footnotes",
  endnotes: "endnotes",
  toc: "toc",
  captions: "captions",
  crossReferences: "cross_references",
};

export async function inspectAgentDocument(doc, toolkit, args, task) {
  const countsOnly = !args.includeDomains?.length;
  const nativeDomains = (args.includeDomains ?? []).filter(
    (domain) => !manualDomains.has(domain),
  );
  const result = await toolkit.dispatch(doc, "superdoc_inspect", {
    countsOnly: countsOnly || nativeDomains.length === 0,
    ...(countsOnly || nativeDomains.length === 0
      ? {}
      : {
          includeDomains: nativeDomains,
          blockOffset: 0,
          blockLimit: 50,
          listLimit: 50,
          tableLimit: 50,
          commentLimit: 50,
          trackedChangeLimit: 50,
        }),
  });
  const manual = {
    footnotes: () => list(doc, "footnotes.list"),
    endnotes: () => doc.footnotes.list({ type: "endnote" }),
    toc: () => list(doc, "toc.list"),
    captions: () => list(doc, "captions.list"),
    crossReferences: () => list(doc, "crossRefs.list"),
  };
  for (const domain of args.includeDomains ?? []) {
    if (!manual[domain]) continue;
    result[domain] = await manual[domain]();
    result.counts = {
      ...(result.counts ?? {}),
      [domain]: result[domain].length,
    };
  }
  const detectedDomains = [
    ...new Set(
      Object.entries(result.counts ?? {})
        .filter(([, count]) => Number(count) > 0)
        .map(([name]) => domainMap[name])
        .filter(Boolean),
    ),
  ];
  const packageInventory = inspectDocxPackage(
    task.input,
    task.maxExpandedBytes,
  );
  const headerFooterSlots = await list(doc, "headerFooters.list");
  const [images, hyperlinks] = await Promise.all([
    list(doc, "images.list"),
    list(doc, "hyperlinks.list"),
  ]);
  const outline = result.outline ?? [];
  const accessibilityIssues = [
    ...images
      .filter(
        (image) =>
          !String(
            image.alt ?? image.description ?? image.properties?.description ?? "",
          ).trim(),
      )
      .map((image) => ({
        code: "IMAGE_ALT_TEXT_MISSING",
        itemId: String(image.sdImageId ?? image.id ?? "image"),
        recovery: "Query the image target and use set_image_alt_text.",
      })),
    ...outline.flatMap((heading, index) => {
      const level = Number(heading.level ?? heading.headingLevel ?? 1);
      const previous = Number(
        outline[index - 1]?.level ?? outline[index - 1]?.headingLevel ?? 0,
      );
      return index > 0 && level > previous + 1
        ? [{
            code: "HEADING_LEVEL_SKIPPED",
            itemId: String(heading.nodeId ?? heading.id ?? index),
            recovery: "Apply a contiguous heading level with apply_style.",
          }]
        : [];
    }),
    ...hyperlinks
      .filter((link) => !String(link.text ?? link.displayText ?? "").trim())
      .map((link, index) => ({
        code: "HYPERLINK_TEXT_EMPTY",
        itemId: String(link.id ?? index),
        recovery: "Replace the empty link with descriptive display text.",
      })),
  ];
  const storyInventory = [
    ...packageInventory.storyInventory,
    ...headerFooterSlots.map((slot) => ({
      type: slot.kind,
      sectionId: slot.section?.sectionId,
      variant: slot.variant ?? "default",
      readable: true,
      story: {
        type: slot.kind,
        sectionId: slot.section?.sectionId,
        variant: slot.variant ?? "default",
      },
      resolution: slot.resolution,
    })),
  ];
  return {
    ...result,
    detectedDomains,
    readingPlan: detectedDomains.map((domain) => ({
      tool: "docx_read",
      arguments:
        domain === "linear"
          ? { representation: "linear", offset: 0, limit: 30 }
          : { representation: "domain", domain, offset: 0, limit: 30 },
    })),
    ...docxActionDiscovery(args),
    accessibilityReport: {
      passed: accessibilityIssues.length === 0,
      issueCount: accessibilityIssues.length,
      checked: ["image-alt-text", "heading-level-order", "hyperlink-display-text"],
      issues: accessibilityIssues,
      notChecked: ["color-contrast", "screen-reader-reading-order", "table-header-semantics"],
    },
    ...packageInventory,
    storyInventory,
  };
}
