import { requireThat } from "../../infrastructure/errors.mjs";

export function storyOf(story) {
  if (!story || story.type === "body")
    return { kind: "story", storyType: "body" };
  if (["header", "footer"].includes(story.type)) {
    requireThat(
      story.sectionId,
      "INVALID_ARGUMENT",
      "Header/footer story selection requires sectionId.",
    );
    return {
      kind: "story",
      storyType: "headerFooterSlot",
      section: { kind: "section", sectionId: story.sectionId },
      headerFooterKind: story.type,
      variant: story.variant ?? "default",
      resolution: "effective",
      onWrite: "error",
    };
  }
  requireThat(
    story.id,
    "INVALID_ARGUMENT",
    `${story.type} story selection requires id.`,
  );
  return {
    kind: "story",
    storyType: story.type,
    [story.type === "textbox" ? "textboxId" : "noteId"]: story.id,
  };
}

export function storyLabel(story) {
  if (!story || story.type === "body") return "body";
  return `${story.type}:${story.sectionId ?? story.id}:${story.variant ?? "default"}`;
}

export function readCoverage(args, domain, nextOffset, additional = {}) {
  const selectedStory = storyLabel(args.story);
  const state = nextOffset === null ? "complete" : "partial";
  const remainingReads = nextOffset === null
    ? []
    : [{
        representation: args.representation ?? "linear",
        ...(args.domain ? { domain: args.domain } : {}),
        ...(args.story ? { story: args.story } : {}),
        ...(args.readSnapshotId
          ? { readSnapshotId: args.readSnapshotId }
          : {}),
        offset: nextOffset,
        limit: args.limit ?? 30,
      }];
  return {
    completeDocument: false,
    windowComplete: nextOffset === null,
    dimensions: {
      sdkVisibleContent: state,
      storyInventory: selectedStory === "body" ? "not_checked" : state,
      annotations: args.includeAnnotations === false ? "not_checked" : state,
      visualLayout: "not_checked",
      unsupportedObjects: "not_checked",
    },
    matrix: {
      [selectedStory]: state,
      visualLayout: "not_checked",
      ...(selectedStory === "body"
        ? {
            headers: "not_checked",
            footers: "not_checked",
            footnotes: "not_checked",
            endnotes: "not_checked",
            textboxes: "not_checked",
          }
        : {}),
    },
    scopes: [{ scope: `${selectedStory}/${domain}`, state }],
    remainingReads,
    notChecked: ["visual-layout", "content-not-exposed-by-SuperDoc"],
    ...additional,
  };
}
