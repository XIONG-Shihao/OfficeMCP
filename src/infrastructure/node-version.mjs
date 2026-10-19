// Runtime gate. OfficeMCP runs on Node or on Bun, whichever the host has.
const MINIMUM_NODE = [24, 12, 0];
const MAXIMUM_NODE_MAJOR = 25;
const MINIMUM_BUN = [1, 3, 13];

function atLeast(value, minimum) {
  const match = /^(\d+)\.(\d+)\.(\d+)/.exec(value ?? "");
  if (!match) return null;
  const version = match.slice(1).map(Number);
  for (let index = 0; index < minimum.length; index += 1) {
    if (version[index] > minimum[index]) return version;
    if (version[index] < minimum[index]) return null;
  }
  return version;
}

export function isSupportedNodeVersion(value) {
  const version = atLeast(value, MINIMUM_NODE);
  return Boolean(version) && version[0] < MAXIMUM_NODE_MAJOR;
}

export function isSupportedBunVersion(value) {
  const version = atLeast(value, MINIMUM_BUN);
  return Boolean(version) && version[0] === MINIMUM_BUN[0];
}

/** The runtime executing this server, as reported to clients and diagnostics. */
export function javascriptRuntime(versions = process.versions) {
  return versions.bun
    ? { kind: "bun", version: versions.bun }
    : { kind: "node", version: versions.node };
}

export function assertSupportedRuntime(versions = process.versions) {
  // Bun also reports a Node compatibility version; under Bun only Bun's own version decides.
  if (versions.bun ? isSupportedBunVersion(versions.bun) : isSupportedNodeVersion(versions.node)) return;
  const runtime = javascriptRuntime(versions);
  const error = new Error(
    `OfficeMCP requires Node >=24.12.0 <25 or Bun >=1.3.13 <2; current runtime is ${runtime.kind} ${runtime.version}.`,
  );
  error.code = "UNSUPPORTED_RUNTIME";
  throw error;
}
