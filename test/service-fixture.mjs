import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { tmpdir } from "node:os";
import { OfficeService } from "../src/application/service.mjs";
import { configFrom } from "../src/infrastructure/files.mjs";

export function createServiceFixture({
  prefix = "officemcp-test-",
  toolset,
  engines = {},
  limits = {},
  prepare,
  payloadOnly = true,
  contractVersion,
} = {}) {
  const root = fs.mkdtempSync(path.join(tmpdir(), prefix));
  const docs = path.join(root, "docs");
  fs.mkdirSync(docs);
  prepare?.(docs);
  const config = path.join(root, "config.json");
  fs.writeFileSync(config, JSON.stringify({
    readRoots: [docs],
    writeRoots: { output: docs },
    stateDirectory: path.join(root, "state"),
    ...(toolset ? { toolset } : {}),
    ...(Object.keys(engines).length ? { engines } : {}),
    ...(Object.keys(limits).length ? { limits } : {}),
  }));
  const service = new OfficeService(configFrom(config));
  return {
    root,
    docs,
    config,
    service,
    call: async (name, args = {}) => {
      const result = await service.call(name, args);
      assert(!result.isError, `${name}: ${JSON.stringify(result.payload)}`);
      if (contractVersion)
        assert.equal(result.payload.contractVersion, contractVersion);
      return payloadOnly ? result.payload : result;
    },
    close() {
      service.close();
      fs.rmSync(root, { recursive: true, force: true });
    },
  };
}
