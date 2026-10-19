import { VERSION, CONTRACT } from "../contracts/registry.mjs";
import { javascriptRuntime } from "../infrastructure/node-version.mjs";
import {
  LEGACY_MCP_PROTOCOL_VERSIONS,
  MCP_PROTOCOL_VERSION,
  TASKS_EXTENSION,
} from "./mcp-protocol.mjs";
import {
  DOCX_AGENT_ACTIONS,
} from "../formats/docx/docx-agent-actions.mjs";

export function serverInfo(service) {
  const agentDocx = ["docx-agent", "full"].includes(service.config.toolset);
  const hasLibreOffice = Boolean(service.config.engines?.libreOffice);
  return {
    data: {
      version: VERSION,
      contractVersion: CONTRACT,
      transport: "stdio",
      protocol: {
        current: MCP_PROTOCOL_VERSION,
        legacyCompatibility: LEGACY_MCP_PROTOCOL_VERSIONS,
        statelessRequests: true,
        applicationState: "explicit handles",
        extensions: [TASKS_EXTENSION],
      },
      javascriptRuntime: javascriptRuntime(),
      platform: `${process.platform}-${process.arch}`,
      operations: service.manifest().map((tool) => tool.name),
      formats: {
        docx: {
          read: agentDocx
            ? "canonical body/story reads, structured hierarchy and typed SuperDoc domains with explicit coverage and remaining reads"
            : "story text and tables plus structured SuperDoc domains; use the docx-agent preset for canonical reads and coverage planning",
          edit: agentDocx
            ? `${DOCX_AGENT_ACTIONS.length} typed high-level actions through SuperDoc v2; request action descriptors from docx_inspect when argument details are needed`
            : "text, formatting and plain tables through SuperDoc v2",
          trackedReview:
            "Tracked text replacement and review actions are available. Unsupported formatted or table-bearing ranges fail explicitly with recovery guidance. docx_compare provides source/candidate comparison.",
          engine: "SuperDoc SDK 2.2.1 / v2 only",
          profiles: ["docx.sdk-only.v1"],
          legacyProfileAliases: [
            "docx.text-replacement.v1",
            "docx.superdoc.v1",
            "docx.story-text.v1",
          ],
          pdfExport: agentDocx
            ? hasLibreOffice
              ? "LibreOffice PDF export is configured; use docx_export to publish a new PDF without overwrite"
              : "docx_export requires configured engines.libreOffice"
            : "Use the docx-agent or full preset for DOCX PDF export",
          layout: agentDocx ? { tool: "docx_layout", available: Boolean(service.config.engines.word || service.config.engines.libreOffice), providers: [service.config.engines.word ? "microsoft-word-macos" : null, hasLibreOffice ? "libreoffice" : null].filter(Boolean), pageTargets: "revision-bound layoutId + page + unique SDK text range", appearance: "explicit Word PDF reference comparison; no universal layout guarantee" } : undefined,
          validation:
            "SDK reopen/readback and hash-chain checks plus independent XML well-formedness, OPC content types, relationship targets, story-part and binary-member preservation evidence; visual fidelity remains separate",
        },
        pdf: {
          read: "page text, bounded search, outlines, annotations, form fields and attachment metadata",
          edit: "page selection, duplication, ordering and rotation; deletion/duplication of referenced pages is refused",
          forms:
            "Fill text, checkbox, dropdown, option-list and radio fields with expected-value guards; XFA and signature-bearing forms are refused",
          ocr: "Optional local Tesseract through office_read with ocr: true; returns recognized words, confidence and bounds. Check office_engine_status for availability.",
          limitations:
            "No general text reflow, annotation editing or semantic table extraction",
        },
        xlsx: {
          read: "xlsx_query: bounded ranges, named tables and business keys; reusable revision-bound targets; stored values, formulas, caches and fresh calculation status",
          edit: "xlsx_apply: atomic typed cells/formulas, keyed rows, formatting, qualified row insertion, fill, sorting/filtering and view settings through a shared preserving writer",
          calculation:
            "local Univer OSS 0.25.2; unsupported dependencies and special formulas remain unresolved. Optional Aspose legacy operations remain separate.",
        },
        csv: {
          read: "xlsx_query or office_read: bounded ranges on sheet CSV, exact literal text and unique business-key lookup across all rows",
          edit: "xlsx_apply or office_edit write_cells: atomic guarded literal cell/keyed updates; source encoding, BOM, unchanged quoting and row endings retained",
          limitations: "CSV has no formula calculation, named tables, formatting or workbook objects. Use XLSX when those representations are required.",
          profile: "csv.literals.v1",
        },
        pptx: {
          read: "slides, shapes, groups, tables, notes, source parts",
          edit: "text, existing speaker notes, shape position and slide ordering; optional Aspose full-shape text and notes",
        },
      },
      limits: service.configLimits(),
      unavailable: [
        "document-rendering",
        ...(!agentDocx ? ["DOCX-page-targets", "visual-comparison"] : []),
        "HTTP",
        "in-place-publication",
        "OS-security-sandbox",
      ],
      retention:
        "Idle document leases expire after artifactRetentionMs; explicit office_cleanup removes expired private artifacts. Receipts expire after operationRetentionMs; permanent operation-ID tombstones prevent reuse. Published outputs and sources are never cleaned.",
    },
  };
}
