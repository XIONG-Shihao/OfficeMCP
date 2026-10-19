/** Frozen optional embedding surface: maintenance fixes only. The integrated app uses the MCP CLI. See docs/SDK.md. */
import { LocalAssetService } from "./application/local-assets.mjs";
import { digest, fingerprint } from "./infrastructure/files.mjs";
import { requireThat } from "./infrastructure/errors.mjs";
/** Host-owned embedding API. Never exposed as arbitrary adapter execution to MCP callers. */
export class OfficeClient {
  constructor({ directory, storage, limits }) {
    requireThat(
      storage &&
        typeof storage.read === "function" &&
        typeof storage.putIfAbsent === "function",
      "INVALID_CONFIG",
      "A storage adapter with read and idempotent putIfAbsent is required.",
    );
    this.storage = storage;
    this.service = new LocalAssetService(directory, limits);
  }
  manifest() {
    return this.service.manifest();
  }
  async call(name, args = {}) {
    const result = await this.service.call(name, args);
    if (name === "office_server_info" && !result.isError) {
      result.payload.data.transport = "embedded";
      result.payload.data.assetTransfer = {
        source: "OfficeClient.importSource",
        publication: "OfficeClient.exportAsset",
      };
    }
    return result;
  }
  async importSource(key, format) {
    const object = await this.storage.read(key);
    requireThat(
      object &&
        Buffer.isBuffer(object.bytes) &&
        digest(object.bytes) === object.revision,
      "INTEGRITY_FAILED",
      "Storage read must return bytes and their exact SHA-256 revision.",
    );
    return this.service.upload(object.bytes, format, object.revision);
  }
  async exportAsset({ assetId, key, operationId }) {
    requireThat(
      typeof operationId === "string" &&
        operationId.length > 0 &&
        operationId.length <= 150 &&
        typeof key === "string" &&
        key.length <= 4096,
      "INVALID_ARGUMENT",
      "Storage publication requires a bounded key and operation ID.",
    );
    const core = this.service.core,
      fp = fingerprint({ assetId, key }),
      old = core.get("storageOperation", operationId, false);
    requireThat(
      !old || old.fingerprint === fp,
      "OPERATION_ID_CONFLICT",
      "Storage operation ID has different arguments.",
    );
    if (old?.result) return old.result;
    core.checkQuota();
    requireThat(
      core.rows("storageOperation").length < 10000,
      "LIMIT_EXCEEDED",
      "Storage receipt quota reached.",
    );
    core.put("storageOperation", operationId, {
      fingerprint: fp,
      status: "running",
    });
    const asset = this.service.download(assetId);
    // Adapter contract requires atomic no-overwrite and repeat safety for the same operation ID.
    const result = await this.storage.putIfAbsent({
      key,
      bytes: asset.bytes,
      operationId,
    });
    requireThat(
      result.revision === asset.hash,
      "INTEGRITY_FAILED",
      "Storage publication revision mismatch.",
    );
    core.put("storageOperation", operationId, {
      fingerprint: fp,
      status: "completed",
      result,
    });
    return result;
  }
  close() {
    this.service.close();
  }
}
