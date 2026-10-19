import fs from "node:fs";
import path from "node:path";
import { digest, fileRevision, fingerprint, newId, regularFile, sourcePath, atomic, snapshotZipFile } from "../infrastructure/files.mjs";
import { decodeCsvBytes } from "#document-core/xlsx/csv-codec.mjs";
import { requireThat } from "../infrastructure/errors.mjs";

const asView = (documentId, revision, candidateId) => ({ documentId, ...(candidateId ? { candidateId } : {}), revision });

/** Owns document/candidate and view-bound handle lifecycles; format adapters supply semantics. */
export class DocumentSessions {
  constructor(service) { this.service = service; }
  open(a, expectedFormat) {
    const service = this.service;
    service.checkQuota(service.config.maxFileBytes);
    requireThat(
      service
        .rows("document")
        .filter(
          (d) =>
            !d.closed &&
            service.now() - d.updatedAt < service.config.artifactRetentionMs,
        ).length < service.config.maxHandles,
      "LIMIT_EXCEEDED",
      "Document handle limit reached.",
    );
    const source = sourcePath(service.config, a.path),
      ext = path.extname(source).toLowerCase(),
      bytes = [".xlsx", ".docx"].includes(ext) ? null : regularFile(source, service.config.maxFileBytes);
    requireThat(
      [".docx", ".pdf", ".xlsx", ".csv", ".pptx"].includes(ext),
      "UNSUPPORTED_FORMAT",
      "Supported documents are DOCX, XLSX, CSV, PPTX and PDF.",
    );
    requireThat(!expectedFormat || ext === "." + expectedFormat, "UNSUPPORTED_FORMAT", `Open a ${expectedFormat?.toUpperCase()} document.`);
    if (ext === ".csv") {
      try { decodeCsvBytes(bytes); }
      catch (error) { requireThat(false, "INVALID_DOCUMENT", error.message); }
    }
    requireThat(
      [".csv", ".xlsx", ".docx"].includes(ext) ? true : ext === ".pdf"
        ? bytes.subarray(0, 5).toString() === "%PDF-"
        : bytes.length >= 4 && bytes.readUInt32LE(0) === 0x04034b50,
      "INVALID_DOCUMENT",
      "File signature does not match the extension.",
    );
    const id = newId("doc"),
      snapshot = service.internal("snapshots", id + ext),
      revision = [".xlsx", ".docx"].includes(ext) ? snapshotZipFile(source,snapshot,service.config.maxFileBytes) : digest(bytes);
    if(bytes)atomic(snapshot, bytes);
    service.put("document", id, {
      source,
      snapshot,
      revision,
      format: ext.slice(1),
      closed: false,
    });
    return {
      document: { id, format: ext.slice(1), sourceRevision: revision },
      data: {
        capabilities: service.manifest().map((tool) => tool.name).filter(
          (name) =>
            name.startsWith(ext.slice(1) + "_") ||
            (ext === ".csv" && ["xlsx_query", "xlsx_apply"].includes(name)) ||
            ([".docx", ".xlsx", ".pptx"].includes(ext) && name.startsWith("office_package_")),
        ),
      },
    };
  }

  async create(a, signal) {
    const service = this.service;
    service.checkQuota(2 * service.config.maxFileBytes);
    requireThat(
      service.rows("document").filter((d) => !d.closed).length <
        service.config.maxHandles,
      "LIMIT_EXCEEDED",
      "Close unused documents before creating another.",
    );
    const importedImages = [];
    try {
      return await service.durable.run("office_create", a, async ({ complete }) => {
      const id = newId("doc"),
        candidateId = newId("candidate"),
        snapshot = service.internal("snapshots", id + "." + a.document.format),
        out = service.internal(
          "candidates",
          candidateId + "." + a.document.format,
        );
      const assets = {};
      const blocks =
        a.document.blocks ?? a.document.slides?.flatMap((s) => s.blocks) ?? [];
      for (const b of blocks)
        if (b.kind === "image") {
          if (b.path) {
            const bytes = regularFile(
              sourcePath(service.config, b.path),
              8 * 1024 * 1024,
            );
            const png = bytes
              .subarray(0, 8)
              .equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
            const jpeg = bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
            requireThat(
              png || jpeg,
              "UNSUPPORTED_FORMAT",
              "Creation images must contain PNG or JPEG bytes.",
            );
            service.checkQuota(bytes.length);
            const file = service.internal("assets", newId("creation-image"));
            fs.writeFileSync(file, bytes, { flag: "wx", mode: 0o600 });
            importedImages.push(file);
            assets[b.path] = file;
            continue;
          }
          const asset = service.get("asset", b.assetId);
          service.doc(asset.documentId);
          requireThat(
            ["image/png", "image/jpeg"].includes(asset.mediaType),
            "UNSUPPORTED_FORMAT",
            "Creation images must be PNG or JPEG.",
          );
          requireThat(
            digest(regularFile(asset.path, 8 * 1024 * 1024)) === asset.hash,
            "INTEGRITY_FAILED",
            "Image asset changed.",
          );
          assets[b.assetId] = asset.path;
        }
      const result = await service.worker(
        "office_create",
        null,
        a,
        { output: snapshot, assets, documentId: id, maxFileBytes: service.config.maxFileBytes },
        signal,
      );
      fs.copyFileSync(snapshot, out, fs.constants.COPYFILE_EXCL);
      service.checkQuota();
      requireThat(
        !signal?.aborted,
        "CANCELLED",
        "Creation cancelled before staging.",
      );
      const receipt = {
        operationId: a.operationId,
        document: { id, format: result.format, sourceRevision: result.revision },
        view: asView(id, result.revision, candidateId),
        data: {
          candidateId,
          candidateRevision: result.revision,
          created: true,
          published: false,
          warnings: result.warnings,
          nextAction: "Inspect with office_read; publish with office_save.",
        },
      };
      service.db.exec("BEGIN IMMEDIATE");
      service.put("document", id, {
        source: snapshot,
        snapshot,
        revision: result.revision,
        format: result.format,
        generated: true,
        closed: false,
      });
      service.put("candidate", candidateId, {
        documentId: id,
        path: out,
        revision: result.revision,
        operations: [],
        published: false,
        discarded: false,
      });
      complete(receipt);
      service.db.exec("COMMIT");
      return receipt;
      });
    } finally {
      for (const file of importedImages) fs.rmSync(file, { force: true });
    }
  }


  transaction(work) {
    const service = this.service;
    service.db.exec("BEGIN IMMEDIATE");
    try { const result = work(); service.db.exec("COMMIT"); return result; }
    catch (error) { if (service.db.isTransaction) service.db.exec("ROLLBACK"); throw error; }
  }

  async fromTemplate(args, format, signal) {
    const service = this.service;
    return service.durable.run(`${format}_create`, args, async ({ complete }) => {
      requireThat(!signal?.aborted, "CANCELLED", "Template creation cancelled.");
      const opened = this.open({ path: args.path }, format);
      const documentId = opened.document.id;
      const document = service.doc(documentId);
      const candidateId = newId("candidate");
      const file = service.internal("candidates", `${candidateId}.${format}`);
      service.checkQuota(fs.statSync(document.snapshot).size);
      fs.copyFileSync(document.snapshot, file, fs.constants.COPYFILE_EXCL);
      const result = { operationId: args.operationId,
        view: asView(documentId, document.revision, candidateId),
        data: { created: true, published: false, sourceTemplateUnchanged: true } };
      return this.transaction(() => {
        service.put("candidate", candidateId, { documentId, path: file, revision: document.revision,
          operations: [], published: false, discarded: false });
        complete(result);
        return result;
      });
    });
  }

  close(documentId, disposition, { view, format } = {}) {
    const service = this.service;
    const document = service.get("document", documentId);
    requireThat(!format || document.format === format, "UNSUPPORTED_FORMAT", `Session is not a ${format?.toUpperCase()}.`);
    if (view && document.closed) return { data: { closed: true } };
    if (view) service.view({ documentId, candidateId: view.candidateId, expectedRevision: view.revision });
    else service.doc(documentId);
    service.docxPool.reset(documentId);
    for (const kind of ["docxTarget", "docxReadSnapshot", "docxLayout"])
      for (const record of service.rows(kind))
        if (record.documentId === documentId) service.delete(kind, record.id);
    for (const candidate of service.rows("candidate"))
      if (candidate.documentId === documentId && !candidate.discarded && !candidate.published && disposition === "discard")
        service.discard(candidate.id, documentId);
    service.put("document", documentId, { ...document, closed: true });
    return { data: { closed: true, disposition } };
  }

  targets(view, result) {
    if (!result.items) return result;
    const convert = ({ nativeTarget, children, ...item }) => {
      const converted = { ...item, ...(children ? { children: children.map(convert) } : {}) };
      if (!nativeTarget) return converted;
      const id = "target_" + fingerprint({ view, nativeTarget }).slice(7);
      this.service.put("docxTarget", id, { documentId: view.documentId, view, locator: nativeTarget });
      return { ...converted, target: { id } };
    };
    return { ...result, items: result.items.map(convert) };
  }

  target(view, target) {
    if (!target) return null;
    const stored = this.service.get("docxTarget", target.id);
    requireThat(fingerprint(stored.view) === fingerprint(view), "TARGET_CONFLICT", "Target belongs to another view. Run docx_query again.");
    return stored.locator;
  }

  readSnapshot(view, identity, id) {
    if (id) {
      const snapshot = this.service.get("docxReadSnapshot", id);
      requireThat(fingerprint(snapshot.identity) === fingerprint(identity), "TARGET_CONFLICT",
        "Read snapshot belongs to another view or representation. Start a new complete read.");
      return id;
    }
    const readSnapshotId = newId("read");
    this.service.put("docxReadSnapshot", readSnapshotId, { documentId: view.documentId, identity });
    return readSnapshotId;
  }

}
