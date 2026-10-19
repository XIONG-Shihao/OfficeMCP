/** Frozen optional embedding surface: maintenance fixes only. The integrated app uses the MCP CLI. See docs/SDK.md. */
import * as fs from "node:fs";
import path from "node:path";
import { OfficeService } from "./service.mjs";
import { configFrom, atomic, digest, newId, regularFile } from "../infrastructure/files.mjs";
import { manifest } from "../contracts/registry.mjs";
import { requireThat } from "../infrastructure/errors.mjs";
export class LocalAssetService {
  constructor(directory, limits = {}) {
    fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
    for (const d of ["input", "output"])
      fs.mkdirSync(path.join(directory, d), { recursive: true, mode: 0o700 });
    const file = path.join(directory, "config.json");
    atomic(
      file,
      Buffer.from(
        JSON.stringify({
          readRoots: [path.join(directory, "input")],
          writeRoots: { output: path.join(directory, "output") },
          stateDirectory: path.join(directory, "state"),
          limits,
        }),
      ),
    );
    this.core = new OfficeService(configFrom(file));
    this.directory = directory;
  }
  manifest() {
    return manifest().map((t) =>
      t.name === "office_open"
        ? {
            ...t,
            description:
              "Open an owned upload asset reference. Use an imported local storage asset.",
            inputSchema: {
              type: "object",
              properties: {
                assetId: { type: "string", minLength: 1, maxLength: 150 },
              },
              required: ["assetId"],
              additionalProperties: false,
            },
          }
        : t,
    );
  }
  upload(bytes, format, expectedHash) {
    this.core.checkQuota();
    requireThat(
      this.core.rows("upload").length < 1000,
      "LIMIT_EXCEEDED",
      "Upload record quota reached.",
    );
    requireThat(
      ["docx", "xlsx", "pptx", "pdf"].includes(format),
      "UNSUPPORTED_FORMAT",
      "Unsupported upload format.",
    );
    requireThat(
      digest(bytes) === expectedHash,
      "INTEGRITY_FAILED",
      "Upload SHA-256 mismatch.",
    );
    requireThat(
      bytes.length <= this.core.config.maxFileBytes,
      "LIMIT_EXCEEDED",
      "Upload too large.",
    );
    let used = 0;
    for (const dir of ["input", "output"])
      for (const entry of fs.readdirSync(path.join(this.directory, dir)))
        used += fs.lstatSync(path.join(this.directory, dir, entry)).size;
    requireThat(
      used + bytes.length <= this.core.config.maxStateBytes,
      "LIMIT_EXCEEDED",
      "Tenant transfer quota reached.",
    );
    const assetId = newId("upload"),
      file = path.join(this.directory, "input", assetId + "." + format);
    atomic(file, bytes);
    this.core.put("upload", assetId, {
      path: file,
      hash: expectedHash,
      format,
    });
    return { assetId, hash: expectedHash, size: bytes.length, format };
  }
  checkTransferQuota(extra = 0) {
    let used = 0;
    for (const dir of ["input", "output"])
      for (const entry of fs.readdirSync(path.join(this.directory, dir)))
        used += fs.lstatSync(path.join(this.directory, dir, entry)).size;
    requireThat(
      used + extra <= this.core.config.maxStateBytes,
      "LIMIT_EXCEEDED",
      "Tenant transfer quota reached.",
    );
  }
  download(assetId) {
    const a = this.core.get("asset", assetId);
    this.core.doc(a.documentId);
    const bytes = regularFile(a.path, this.core.config.maxFileBytes);
    requireThat(
      digest(bytes) === a.hash,
      "INTEGRITY_FAILED",
      "Artifact changed.",
    );
    return { bytes, mediaType: a.mediaType, hash: a.hash };
  }
  async call(name, args = {}) {
    try {
      if (name === "office_open") {
        requireThat(
          Object.keys(args).length === 1 && typeof args.assetId === "string",
          "INVALID_ARGUMENT",
          "Use only an owned assetId.",
        );
        const a = this.core.get("upload", args.assetId);
        requireThat(
          Date.now() - a.updatedAt < this.core.config.artifactRetentionMs,
          "HANDLE_EXPIRED",
          "Upload expired.",
        );
        requireThat(
          digest(regularFile(a.path, this.core.config.maxFileBytes)) === a.hash,
          "INTEGRITY_FAILED",
          "Upload changed.",
        );
        args = { path: a.path };
      }
      if (
        ["office_commit", "office_save"].includes(name) &&
        !this.core.get("operation", args.operationId, false)
      )
        this.checkTransferQuota(this.core.config.maxFileBytes);
      const result = await this.core.call(name, args);
      if (name === "office_cleanup" && !result.isError)
        for (const a of this.core.rows("upload"))
          if (
            Date.now() - a.updatedAt >= this.core.config.artifactRetentionMs &&
            !this.core.rows("document").some((d) => d.source === a.path)
          ) {
            fs.unlinkSync(a.path);
            this.core.delete("upload", a.id);
          }
      // Same receipts and validation semantics; deployment paths are private implementation details.
      const sanitize = (v) =>
        Array.isArray(v)
          ? v.map(sanitize)
          : v && typeof v === "object"
            ? Object.fromEntries(
                Object.entries(v)
                  .filter(([k]) => k !== "path")
                  .map(([k, x]) => [k, sanitize(x)]),
              )
            : v;
      const sanitized = sanitize(result);
      return sanitized;
    } catch (e) {
      return this.core.failure(e, args);
    }
  }
  close() {
    this.core.close();
  }
}
