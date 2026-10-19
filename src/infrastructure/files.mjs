import * as fs from "node:fs";
import path from "node:path";
import { createHash, randomUUID } from "node:crypto";
import { OfficeError, requireThat } from "./errors.mjs";
export const digest = (bytes) =>
  "sha256:" + createHash("sha256").update(bytes).digest("hex");
export function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object")
    return Object.fromEntries(
      Object.keys(value)
        .sort()
        .map((k) => [k, canonical(value[k])]),
    );
  return value;
}
export const fingerprint = (value) => digest(JSON.stringify(canonical(value)));
export const newId = (prefix) => `${prefix}_${randomUUID()}`;
export function confined(root, file) {
  const r = path.relative(root, file);
  return (
    r === "" ||
    (!path.isAbsolute(r) && r !== ".." && !r.startsWith(".." + path.sep))
  );
}
export function regularFile(file, maxBytes) {
  const fd = fs.openSync(
    file,
    fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW ?? 0),
  );
  try {
    const st = fs.fstatSync(fd);
    requireThat(
      st.isFile() && st.size <= maxBytes,
      "LIMIT_EXCEEDED",
      "Input must be a regular file within the configured size limit.",
    );
    const bytes = fs.readFileSync(fd);
    requireThat(
      bytes.length <= maxBytes,
      "LIMIT_EXCEEDED",
      "Input grew beyond the limit.",
    );
    return bytes;
  } finally {
    fs.closeSync(fd);
  }
}
/** Hash large snapshots without allocating another whole-document Buffer. */
export function fileRevision(file, maxBytes) {
  const fd = fs.openSync(file, fs.constants.O_RDONLY | (fs.constants.O_NOFOLLOW ?? 0));
  try {
    const stat = fs.fstatSync(fd);
    requireThat(stat.isFile() && stat.size <= maxBytes, "LIMIT_EXCEEDED",
      "Input must be a regular file within the configured size limit.");
    const hash = createHash("sha256");
    const buffer = Buffer.allocUnsafe(Math.min(1024 * 1024, Math.max(1, stat.size)));
    let total = 0, count;
    while ((count = fs.readSync(fd, buffer, 0, buffer.length, null)) > 0) {
      total += count;
      requireThat(total <= maxBytes, "LIMIT_EXCEEDED", "Input grew beyond the limit.");
      hash.update(buffer.subarray(0, count));
    }
    requireThat(total === stat.size, "INTEGRITY_FAILED", "File size changed while reading.");
    return "sha256:" + hash.digest("hex");
  } finally { fs.closeSync(fd); }
}
export function atomic(file, bytes) {
  const temp = file + "." + randomUUID() + ".tmp";
  const fd = fs.openSync(temp, "wx", 0o600);
  try {
    fs.writeFileSync(fd, bytes);
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
  fs.renameSync(temp, file);
  syncDirectory(path.dirname(file));
}
export function syncDirectory(dir) {
  const fd = fs.openSync(dir, "r");
  try {
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
}
export function configFrom(file) {
  const cfg = JSON.parse(fs.readFileSync(file, "utf8"));
  const base = path.dirname(path.resolve(file));
  if (["docx", "docx-advanced"].includes(cfg.toolset))
    cfg.toolset = "docx-agent";
  requireThat(
    cfg.toolset === undefined ||
      ["all", "recommended", "docx-agent", "full"].includes(cfg.toolset),
    "INVALID_CONFIG",
    "toolset must be all, recommended, docx-agent or full.",
  );
  requireThat(
    Array.isArray(cfg.readRoots) &&
      cfg.readRoots.length > 0 &&
      cfg.writeRoots &&
      typeof cfg.writeRoots === "object" &&
      !Array.isArray(cfg.writeRoots) &&
      typeof cfg.stateDirectory === "string",
    "INVALID_CONFIG",
    "Configure readRoots, writeRoots and stateDirectory.",
  );
  const readRoots = cfg.readRoots.map((r) =>
    fs.realpathSync(path.resolve(base, r)),
  );
  const writeRoots = Object.fromEntries(
    Object.entries(cfg.writeRoots).map(([k, v]) => [
      k,
      fs.realpathSync(path.resolve(base, v)),
    ]),
  );
  const state = path.resolve(base, cfg.stateDirectory);
  fs.mkdirSync(state, { recursive: true, mode: 0o700 });
  requireThat(
    !fs.lstatSync(state).isSymbolicLink(),
    "INVALID_CONFIG",
    "State directory cannot be a symlink.",
  );
  fs.chmodSync(state, 0o700);
  const stateDirectory = fs.realpathSync(state);
  requireThat(
    [...readRoots, ...Object.values(writeRoots)].every(
      (r) =>
        fs.statSync(r).isDirectory() &&
        !confined(r, stateDirectory) &&
        !confined(stateDirectory, r),
    ),
    "INVALID_CONFIG",
    "Keep private state separate from read/write roots.",
  );
  return {
    readRoots,
    writeRoots,
    stateDirectory,
    toolset: cfg.toolset ?? "all",
    engines: engineConfig(cfg.engines, base),
    maxFileBytes: 320 * 1024 * 1024,
    maxExpandedBytes: 384 * 1024 * 1024,
    maxStateBytes: 2 * 1024 * 1024 * 1024,
    maxHandles: 64,
    maxActiveOperations: 10000,
    maxDocxActors: 2,
    maxDocxResidentActors: 2,
    maxDocxQueuedRequests: 32,
    maxDocxProcessBytes: 8 * 1024 ** 3,
    docxIdleMs: 60000,
    artifactRetentionMs: 7 * 86400000,
    operationRetentionMs: 30 * 86400000,
    workerTimeoutMs: 120000,
    maxResponseBytes: 200000,
    ...Object.fromEntries(
      Object.entries(cfg.limits ?? {}).map(([k, v]) => {
        requireThat(
          [
            "maxFileBytes",
            "maxExpandedBytes",
            "maxStateBytes",
            "maxHandles",
            "maxActiveOperations",
            "maxDocxActors",
            "maxDocxResidentActors",
            "maxDocxQueuedRequests",
            "maxDocxProcessBytes",
            "docxIdleMs",
            "artifactRetentionMs",
            "operationRetentionMs",
            "workerTimeoutMs",
            "maxResponseBytes",
          ].includes(k) &&
            Number.isSafeInteger(v) &&
            v > 0 &&
            (k !== "maxDocxActors" || v <= 4) &&
            (k !== "maxDocxResidentActors" || v <= 4) &&
            (k !== "maxDocxQueuedRequests" || v <= 128),
          "INVALID_CONFIG",
          "Invalid limit.",
        );
        return [k, v];
      }),
    ),
  };
}
function engineConfig(value = {}, base) {
  requireThat(
    value &&
      typeof value === "object" &&
      !Array.isArray(value) &&
      Object.keys(value).every((k) =>
        [
          "python",
          "asposeLicense",
          "asposeCellsLicense",
          "asposeSlidesLicense",
          "libreOffice",
          "word",
          "wordScratch",
          "tesseract",
          "allowEvaluation",
        ].includes(k),
      ),
    "INVALID_CONFIG",
    "Unknown engine setting.",
  );
  requireThat(
    value.allowEvaluation === undefined ||
      typeof value.allowEvaluation === "boolean",
    "INVALID_CONFIG",
    "allowEvaluation must be boolean.",
  );
  const result = { allowEvaluation: value.allowEvaluation === true };
  for (const key of [
    "python",
    "asposeLicense",
    "asposeCellsLicense",
    "asposeSlidesLicense",
    "libreOffice",
    "word",
    "wordScratch",
    "tesseract",
  ])
    if (value[key] !== undefined) {
      requireThat(
        typeof value[key] === "string" && value[key].length > 0,
        "INVALID_CONFIG",
        "Engine paths must be nonempty strings.",
      );
      // A venv executable is normally a symlink; resolving it loses its venv.
      if (key === 'wordScratch') {
        result[key] = path.resolve(base, value[key]);
        requireThat(!fs.existsSync(result[key]) || fs.statSync(result[key]).isDirectory(), 'INVALID_CONFIG', 'Word scratch must be a private directory.');
        continue;
      }
      result[key] =
        key === "python"
          ? path.resolve(base, value[key])
          : fs.realpathSync(path.resolve(base, value[key]));
      requireThat(
        key === 'word' ? process.platform === 'darwin' && fs.statSync(result[key]).isDirectory() && fs.existsSync(path.join(result[key], 'Contents', 'MacOS', 'Microsoft Word')) : fs.statSync(result[key]).isFile(),
        "INVALID_CONFIG",
        key === 'word' ? 'Word must be an installed Microsoft Word.app bundle on macOS.' : "Engine path must be a regular file.",
      );
    }
  return result;
}
export function sourcePath(config, input) {
  requireThat(
    path.isAbsolute(input),
    "ACCESS_DENIED",
    "Use an absolute path inside a configured read root.",
  );
  let real;
  try {
    real = fs.realpathSync(input);
  } catch {
    throw new OfficeError("NOT_FOUND", "Source file is unavailable.");
  }
  requireThat(
    config.readRoots.some((r) => confined(r, real)),
    "ACCESS_DENIED",
    "Source is outside configured read roots.",
  );
  return real;
}
export function destinationPath(config, target) {
  const root = config.writeRoots[target.rootId];
  requireThat(
    root && !path.isAbsolute(target.relativePath),
    "ACCESS_DENIED",
    "Unknown output root or absolute destination.",
  );
  const out = path.resolve(root, target.relativePath);
  requireThat(
    confined(root, out) && out !== root,
    "ACCESS_DENIED",
    "Destination escapes output root.",
  );
  let parent;
  try {
    parent = fs.realpathSync(path.dirname(out));
  } catch {
    throw new OfficeError(
      "NOT_FOUND",
      "Destination parent must already exist.",
    );
  }
  requireThat(
    confined(root, parent),
    "ACCESS_DENIED",
    "Destination parent escapes output root.",
  );
  return path.join(parent, path.basename(out));
}

/** Copy and hash an authorized ZIP document snapshot with fixed-size buffers. */
export function snapshotZipFile(source,target,maxBytes){
  const input=fs.openSync(source,fs.constants.O_RDONLY|(fs.constants.O_NOFOLLOW??0));
  const temporary=target+'.'+randomUUID()+'.tmp';let output;
  try{const before=fs.fstatSync(input);requireThat(before.isFile()&&before.size<=maxBytes,'LIMIT_EXCEEDED','Input must be a regular ZIP document within the configured size limit.');
    requireThat(before.size>=4,'INVALID_DOCUMENT','File signature does not match the extension.');output=fs.openSync(temporary,'wx',0o600);const hash=createHash('sha256'),buffer=Buffer.allocUnsafe(Math.min(1024*1024,before.size));let total=0,count;
    while((count=fs.readSync(input,buffer,0,buffer.length,null))>0){if(total===0)requireThat(count>=4&&buffer.readUInt32LE(0)===0x04034b50,'INVALID_DOCUMENT','File signature does not match the extension.');total+=count;requireThat(total<=maxBytes,'LIMIT_EXCEEDED','Input grew beyond the limit.');const chunk=buffer.subarray(0,count);hash.update(chunk);fs.writeFileSync(output,chunk);}
    const after=fs.fstatSync(input);requireThat(total===before.size&&after.size===before.size&&after.mtimeMs===before.mtimeMs&&after.ctimeMs===before.ctimeMs,'INTEGRITY_FAILED','File changed while snapshotting.');fs.fsyncSync(output);fs.closeSync(output);output=undefined;fs.renameSync(temporary,target);syncDirectory(path.dirname(target));return 'sha256:'+hash.digest('hex');
  }finally{fs.closeSync(input);if(output!==undefined)fs.closeSync(output);fs.rmSync(temporary,{force:true});}
}
export const snapshotXlsxFile = snapshotZipFile;
