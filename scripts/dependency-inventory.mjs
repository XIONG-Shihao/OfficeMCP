import { readFileSync, writeFileSync } from "node:fs";
const lock = JSON.parse(
  readFileSync(new URL("../npm-shrinkwrap.json", import.meta.url)),
);
const dependencies = Object.entries(lock.packages)
  .filter(([p, info]) => p && !info.dev)
  .map(([location, p]) => ({
    location,
    name: p.name ?? location.split("node_modules/").at(-1),
    version: p.version,
    integrity: p.integrity ?? null,
    license: p.license ?? "not declared in lockfile",
    optional: p.optional ?? false,
    os: p.os ?? null,
    cpu: p.cpu ?? null,
  }))
  .sort((a, b) => a.location.localeCompare(b.location));
writeFileSync(
  new URL("../contracts/dependencies.json", import.meta.url),
  JSON.stringify(
    {
      format: "officemcp.lock-inventory.v1",
      package: lock.name,
      version: lock.version,
      source: "npm-shrinkwrap.json",
      dependencies,
    },
    null,
    2,
  ) + "\n",
);
console.log(
  `Recorded ${dependencies.length} locked runtime packages, including optional platform variants.`,
);
