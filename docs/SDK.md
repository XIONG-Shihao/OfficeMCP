> Frozen optional embedding API: maintenance and correctness fixes only. No new SDK/storage features or extraction package are planned. The integrated app launches the MCP CLI.

# Embedding and storage adapters

The optional Node ESM API uses the same `OfficeService`, schemas, worker adapters,
validation, operation IDs and publication rules as MCP. It contains no app imports. The JavaScript class/storage interfaces remain available; `call` and `manifest` follow the current MCP contract, so the removed rendering tools are not retained through this API.

```js
import { OfficeClient } from 'officemcp-local-preview';
import { SQLiteObjectStore } from 'officemcp-local-preview/storage';

const storage = new SQLiteObjectStore('/absolute/objects.sqlite');
const office = new OfficeClient({
  directory: '/absolute/private-office-workspace',
  storage,
});
try {
  const source = await office.importSource('existing-object-key', 'docx');
  const result = await office.call('office_open', { assetId: source.assetId });
  // Read/edit/validate/commit with the published tool schemas.
  // After commit, export its assetId to a new object-store key:
  // await office.exportAsset({ assetId, key: 'reviewed.docx', operationId: 'export-1' });
} finally {
  office.close();
  storage.close();
}
```

`call` returns the same `{isError?, payload}` service result. `manifest()` exposes
the asset-reference input used by this embedding profile. Type declarations are
included in the package; Node TypeScript consumers should install their usual
`@types/node` development dependency.

A host-owned storage adapter implements:

- `read(key)` → `{bytes: Buffer, revision: 'sha256:...'}`.
- `putIfAbsent({key, bytes, operationId})` → `{key, revision, operationId, published:true}`.

`putIfAbsent` must atomically refuse an existing destination, persist bytes and
its operation identity together, and return the same receipt when identical work
is retried after a crash or lost reply. Changed arguments for the same ID must fail.
OfficeMCP checks revision hashes and records publication in the same durable
operation state machine used by create, edit, save and export.
It cannot manufacture atomicity for an adapter that violates this contract.

`SQLiteObjectStore` is the included persistent reference adapter: object BLOBs and
operation IDs share one transaction, with WAL, FULL synchronization and quotas.
It supports multiple local database connections using SQLite serialization. It
is not a cloud object-store or distributed coordinator. Keys are opaque strings,
not paths. These object-store limits are separate from the OfficeService workspace limits. Limits default to 20 MiB per object and 256 MiB of stored object bytes.
A 10,000-object count limit also bounds zero-byte objects (configurable as
`maxObjects`). There is no deletion API, so operation IDs and publication receipts remain reserved.
The embedding workspace retains at most 10,000 export receipts and fails closed
when full. The caller owns storage cleanup/archival and lifecycle.

Embedding and external storage are optional. The local stdio executable does not
load caller-provided adapter code from tool arguments or expose arbitrary engine
execution. Optional local PDF-conversion/OCR/calculation engines are configured by the host;
see [ENGINES.md](ENGINES.md) for their availability and qualification limits.

The installed-package test imports these public exports from an empty directory,
imports a DOCX from SQLite, runs edit → validate → commit, exports its bytes, and
checks the same preservation and no-overwrite semantics used over MCP. Restart
and lost-reply behavior are covered by the SDK/storage suite.
