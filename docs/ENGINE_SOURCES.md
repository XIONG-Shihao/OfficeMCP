# External engine source setup

The OfficeMCP Git repository contains no DOCX/XLSX engine source or editor UI.
Build adapters consume these external component directories:

| Setting | Required directory and inputs |
| --- | --- |
| `OFFICEMCP_XLSX_ENGINE_SOURCE` | XLSX component root: `core/`, `native/` (Cargo sources, lockfile and assets), `scripts/build-native.mjs`, `scripts/export-core.mjs`, `scripts/export-native.mjs`. Tests additionally use `fixtures/`, `scripts/test-support.mjs`, `scripts/test/` and `scripts/support/process-memory.mjs`. |
| `OFFICEMCP_DOCX_ENGINE_SOURCE` | DOCX component root: `core/`, `scripts/export-core.mjs`, `scripts/prepare-observer.mjs`. Tests additionally use `fixtures/generated.mjs`. |

Currently both components are available from
`git@github.com:XIONG-Shihao/Agent_Platform.git`, under `editor/xlsx` and
`editor/docx`. Repository access is required separately. Reference source revision:
`b9488bb33285b4e7ea5f216a44bf9549aa907816`.
This records the source used for this extraction; changing the engine revision
requires rerunning engine identity and document workflow checks.

## Checkout and build

Use separate directories. A sparse checkout retrieves the component folders
without checking out the main app or its integration code:

```sh
git clone git@github.com:XIONG-Shihao/OfficeMCP.git OfficeMCP
git clone --filter=blob:none --sparse git@github.com:XIONG-Shihao/Agent_Platform.git office-engine-sources
git -C office-engine-sources sparse-checkout set editor/xlsx editor/docx
git -C office-engine-sources checkout b9488bb33285b4e7ea5f216a44bf9549aa907816
export OFFICEMCP_XLSX_ENGINE_SOURCE="$PWD/office-engine-sources/editor/xlsx"
export OFFICEMCP_DOCX_ENGINE_SOURCE="$PWD/office-engine-sources/editor/docx"
cd OfficeMCP
npm ci --ignore-scripts
npm run build
npm test
```

Use Node `>=24.12.0 <25` for these npm commands, a current Rust toolchain supporting
edition 2024, and on macOS Xcode Command Line Tools (`clang`). Cargo uses the
component's committed lockfile; its first build needs registry/cache access.
OfficeMCP installs the JavaScript runtime dependencies from its own shrinkwrap;
the external component folders do not require an editor UI build or npm install
for engine preparation. Paths may contain spaces.

In the original workspace layout, `OfficeMCP/` and `editor/` are siblings, so the
adapters retain that default. For independent checkouts set both absolute paths
as above. `OFFICEMCP_DOCX_SOURCE` used by workflow scripts means the **OfficeMCP
package under test**, not the external DOCX engine; do not confuse the settings.

## Prepared runtime versus source

`npm run build`, `npm test` and `npm pack` prepare the engines through the same
adapter. The source-check tests compare generated JavaScript and native engine
bytes with the selected external inputs. `node scripts/prepare-core.mjs --check`
checks an existing preparation without rebuilding it.

`.generated/` and native build outputs are ignored in Git. Runtime packages include
headless exports and the helper for the build machine's platform/architecture;
end users of an assembled package do not need these external source folders or
Rust. Build separately and qualify for each target. macOS arm64 is exercised;
this repository does not claim other platforms are verified.

External component checkout/update, engine licenses and engine source ownership
remain separate from OfficeMCP. No application shell, app settings, harness,
workspace integration or desktop package command is needed for the MCP source build.
