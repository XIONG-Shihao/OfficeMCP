import { resultSchemas } from "../../contracts/results.mjs";
import { z } from "zod";
export const docxDomains = z.enum([
  "paragraphs",
  "sections",
  "headers_footers",
  "fields",
  "footnotes",
  "hyperlinks",
  "images",
  "tables",
  "tracked_changes",
  "comments",
  "content_controls",
  "lists",
  "bookmarks",
  "capabilities",
]);
export function reviewTools(view) {
  return {
    docx_read_details: {
      description: "Read structured DOCX details through SuperDoc v2: section/header relationships, fields, footnotes, images, tables and other domains. Capability entries describe the SDK, not exposed MCP operations.",
      input: z.object({
        view,
        domain: docxDomains,
        offset: z.number().int().min(0).default(0),
        limit: z.number().int().min(1).max(100).default(30),
      }),
      output: resultSchemas.page,
      risk: "read",
      handler: "workflow.docx-details",
      execute: (service, name, args, signal) => service.commands.readDocxDetails(name, args, signal),
      execution: "taskEligible",
      engines: ["superdoc-v2"],
      presets: ["all"],
    },
  };
}
