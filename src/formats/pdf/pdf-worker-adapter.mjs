import * as mupdf from "mupdf";
import { requireThat } from "../../infrastructure/errors.mjs";

const omissions = ["OCR", "reading-order-inference", "non-text-content"];

function lines(page) {
  const text = page.toStructuredText("preserve-whitespace");
  try {
    return (JSON.parse(text.asJSON()).blocks ?? []).flatMap((block) =>
      (block.lines ?? []).map((line) => ({
        text: line.text ?? "",
        transform: [1, 0, 0, 1, line.x ?? line.bbox?.x ?? 0, line.y ?? line.bbox?.y ?? 0],
        width: line.bbox?.w ?? 0,
        height: line.bbox?.h ?? 0,
        hasEOL: true,
      })),
    );
  } finally {
    text.destroy();
  }
}

function outlineItems(nodes, depth = 0, items = []) {
  requireThat(items.length < 10000 && depth < 100, "LIMIT_EXCEEDED", "Outline exceeds inspection limits.");
  for (const node of nodes ?? []) {
    items.push({
      title: node.title,
      depth,
      destination: node.page == null ? null : String(node.page),
      url: node.uri ?? null,
      bold: Boolean(node.bold),
      italic: Boolean(node.italic),
    });
    outlineItems(node.down, depth + 1, items);
  }
  return items;
}

function search(task, doc) {
  const { pageFrom, pageTo, query, caseSensitive } = task.args;
  requireThat(
    pageTo >= pageFrom && pageTo <= doc.countPages() && pageTo - pageFrom < 50,
    "INVALID_ARGUMENT",
    "Choose an existing range of at most 50 pages.",
  );
  const needle = caseSensitive ? query : query.toLocaleLowerCase();
  const items = [];
  for (let pageNumber = pageFrom; pageNumber <= pageTo; pageNumber++) {
    const page = doc.loadPage(pageNumber - 1);
    try {
      for (const line of lines(page)) {
        const source = caseSensitive ? line.text : line.text.toLocaleLowerCase();
        let offset = 0;
        while ((offset = source.indexOf(needle, offset)) >= 0) {
          requireThat(items.length < 10000, "LIMIT_EXCEEDED", "Too many matches; narrow page range or query.");
          items.push({
            page: pageNumber,
            match: line.text.slice(offset, offset + query.length),
            context: line.text.slice(Math.max(0, offset - 100), offset + query.length + 100),
            offset,
            positions: [{ transform: line.transform, width: line.width, height: line.height }],
            coordinateSpace: "top-left page points",
          });
          offset += Math.max(1, needle.length);
        }
      }
    } finally {
      page.destroy();
    }
  }
  return { items, domain: "page-range-search", pageFrom, pageTo, omissions };
}

export async function runPdfWorker(task, bytes) {
  const doc = new mupdf.PDFDocument(bytes);
  doc.disableJS();
  try {
    requireThat(doc.countPages() <= 2000, "LIMIT_EXCEEDED", "PDF exceeds page limit.");
    if (task.kind === "pdf_inspect")
      return {
        pageCount: doc.countPages(),
        supportedDomains: ["page-text"],
        omissions: ["annotations", "forms", "attachments", "OCR", "tagged-structure"],
      };
    if (task.kind === "pdf_outline")
      return { items: outlineItems(doc.loadOutline()), domain: "outline", omissions: ["destination-resolution", "link-content"] };
    if (task.kind === "pdf_search") return search(task, doc);
    requireThat(task.kind === "pdf_read", "UNSUPPORTED_FEATURE", "Unknown PDF read operation.");
    requireThat(task.args.page <= doc.countPages(), "TARGET_CONFLICT", "PDF page is unavailable.");
    const page = doc.loadPage(task.args.page - 1);
    try {
      return { domain: "page-text", page: task.args.page, items: lines(page), omissions };
    } finally {
      page.destroy();
    }
  } finally {
    doc.destroy();
  }
}
