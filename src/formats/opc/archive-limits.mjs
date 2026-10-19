import { requireThat } from "../../infrastructure/errors.mjs";
// Container metadata only: no decompression or document parsing. Uses Node Buffer.
export function checkArchiveLimits(bytes, maxExpandedBytes) {
  // Inspect the central directory before allocating expanded member buffers.
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--)
    if (bytes.readUInt32LE(i) === 0x06054b50) {
      end = i;
      break;
    }
  requireThat(
    end >= 0,
    "INVALID_DOCUMENT",
    "Input must be a supported OOXML ZIP package.",
  );
  const entries = bytes.readUInt16LE(end + 10),
    start = bytes.readUInt32LE(end + 16);
  requireThat(
    bytes.readUInt16LE(end + 4) === 0 &&
      bytes.readUInt16LE(end + 6) === 0 &&
      entries < 5000 &&
      start < end,
    "UNSUPPORTED_FEATURE",
    "Multipart, ZIP64 or oversized ZIP archives are unsupported.",
  );
  let at = start,
    total = 0;
  const names = new Set();
  for (let i = 0; i < entries; i++) {
    requireThat(
      at + 46 <= end && bytes.readUInt32LE(at) === 0x02014b50,
      "INVALID_DOCUMENT",
      "Invalid ZIP directory.",
    );
    const size = bytes.readUInt32LE(at + 24),
      n = bytes.readUInt16LE(at + 28),
      extra = bytes.readUInt16LE(at + 30),
      comment = bytes.readUInt16LE(at + 32);
    const name = bytes.subarray(at + 46, at + 46 + n).toString("utf8");
    requireThat(
      !(bytes.readUInt16LE(at + 8) & 1) &&
        !name.startsWith("/") &&
        !name.includes("\\") &&
        !name.split("/").includes("..") &&
        !names.has(name),
      "INVALID_DOCUMENT",
      "Unsafe, encrypted or duplicate ZIP member.",
    );
    total += size;
    requireThat(
      total <= maxExpandedBytes && size !== 0xffffffff,
      "LIMIT_EXCEEDED",
      "ZIP expansion exceeds limit.",
    );
    names.add(name);
    at += 46 + n + extra + comment;
  }
  requireThat(at <= end, "INVALID_DOCUMENT", "Invalid ZIP directory length.");
  return { entries, expandedBytes: total };
}
