# Repository boundary

Remote: `git@github.com:XIONG-Shihao/OfficeMCP.git`.

This repository contains the portable local MCP product: CLI, protocol, document
services, external engine build adapters, schemas, public SDK, tests and product
documentation. It contains no full desktop application, editor UI, Tauri/Electron
shell, app account/settings service or application integration adapter.

Engine source checkouts stay outside this repository; see
[Engine source setup](ENGINE_SOURCES.md). Generated engine exports, native executables, npm dependencies, desktop bundles,
packaged archives, caches, local documents and evaluation outputs are ignored.
Optional document engines and their licenses are supplied by the operator.
The package remains private and `UNLICENSED`; pushing source does not grant an
open-source license to it or change third-party licenses.

When this repository lives at `OfficeMCP/` inside an application checkout, it has
its own `.git` and remote. Run component Git commands from this directory or use
`git -C OfficeMCP ...`. The parent repository still tracks its existing copy;
this setup does not silently convert that directory to a submodule or commit
unrelated application work. Component commits and parent commits are separate.
A consuming application must explicitly choose a committed component revision.

Some old integration documents and saved-account benchmarks may remain locally
in the parent checkout. Their exact paths are excluded in this repository's
`.gitignore`; they are not portable product inputs. The portable live-agent
runner uses caller-supplied credentials, never a saved application account.

Use [Release verification](RELEASE.md) for product checks and
[Usage](USAGE.md) for public stdio configuration. Full app builds and combined UI
verification belong to the application repository.
