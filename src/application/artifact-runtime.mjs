import fs from "node:fs";
import { digest, newId, regularFile } from "../infrastructure/files.mjs";
import { requireThat } from "../infrastructure/errors.mjs";

const imageTypes = ["image/png", "image/jpeg"];
const imageLimit = 8 * 1024 * 1024;

/** Allocation, worker output ownership and native-image transport for every format. */
export class ArtifactRuntime {
  constructor(service) { this.service = service; }

  allocate(prefix, extension = ".bin", reservation = 0, directory = "assets") {
    this.service.checkQuota(reservation);
    return this.service.internal(directory, newId(prefix) + extension);
  }

  read(assetId, { mediaTypes, maxBytes = this.service.config.maxFileBytes } = {}) {
    const asset = this.service.get("asset", assetId);
    this.service.doc(asset.documentId);
    requireThat(!mediaTypes || mediaTypes.includes(asset.mediaType), "UNSUPPORTED_FORMAT", "Asset media type is not supported by this operation.");
    const bytes = regularFile(asset.path, maxBytes);
    requireThat(digest(bytes) === asset.hash, "INTEGRITY_FAILED", "Artifact changed.");
    return { asset, bytes };
  }

  image(assetId, mediaTypes = imageTypes) {
    const { asset, bytes } = this.read(assetId, { mediaTypes, maxBytes: imageLimit });
    return { asset, bytes, nativeImage: { type: "image", mimeType: asset.mediaType, data: bytes.toString("base64") } };
  }

  async produce({ kind, input, args = {}, context = {}, documentId,
    prefix = "artifact", extension = ".bin", reservation = 0, output,
    mediaType, includeImage = false, requireImage = false }, signal) {
    const service = this.service;
    const file = output ?? this.allocate(prefix, extension, reservation);
    let accepted = false, assetId;
    try {
      const result = await service.worker(kind, input, args, { ...context, output: file }, signal);
      requireThat(!signal?.aborted, "CANCELLED", "Artifact creation cancelled before acceptance.");
      service.doc(documentId);
      const type = mediaType ?? result.mediaType;
      if (requireImage) requireThat(imageTypes.includes(type), "UNSUPPORTED_FORMAT", "The selected image is not PNG or JPEG.");
      assetId = service.asset(file, documentId, type);
      const wantsImage = includeImage === "auto"
        ? imageTypes.includes(type) && fs.statSync(file).size <= imageLimit
        : includeImage;
      const response = { data: { ...result, assetId },
        ...(wantsImage ? { nativeImage: this.image(assetId).nativeImage } : {}) };
      accepted = true;
      return response;
    } finally {
      if (!accepted) {
        if (assetId) service.delete("asset", assetId);
        try { fs.unlinkSync(file); } catch {}
      }
    }
  }

}
