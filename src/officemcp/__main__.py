"""Command line: `serve --config FILE` runs the MCP server on stdio; `contracts [--check]` writes the generated files."""

import argparse
import sys

import anyio


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="officemcp", description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    serve = commands.add_parser("serve", help="Run the MCP server on stdio")
    serve.add_argument("--config", required=True, help="JSON with roots, engines and limits")
    contracts = commands.add_parser("contracts", help="Write contracts/catalog.json and agent/OfficeMCP.md")
    contracts.add_argument("--check", action="store_true", help="Fail if they are out of date instead of writing")
    args = parser.parse_args(argv)
    if args.command == "serve":
        from .files import Config
        from .server import serve as run

        anyio.run(run, Config.load(args.config))
        return 0
    from .contracts import write

    stale = write(check=args.check)
    if args.check and stale:
        print("Out of date; run python -m officemcp contracts: " + ", ".join(stale), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
