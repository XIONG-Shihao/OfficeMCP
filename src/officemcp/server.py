"""The MCP server: the seven tools over stdio. Each call runs the matching Service method."""

import anyio
import json

import mcp.types as types
from mcp.server import Server
from mcp.server.stdio import stdio_server

from .files import Config
from .service import Service
from .tools import RESULT, TOOLS

VERSION = "0.1.0"
INSTRUCTIONS = (
    "Work on PDF, DOCX, XLSX and PPTX files with standardized actions. Typical flow: office_inspect a file "
    "(note its revision and addresses), office_read or office_find what you need, then office_edit with "
    "actions from office_actions and the revision. Edits save the file in place (or as a new file); they refuse "
    "to touch text that is not what you expected, and to lose content unless you pass allowLoss."
)


def dispatch(service: Service, name: str, arguments: dict) -> dict:
    calls = {
        "office_actions": service.actions,
        "office_inspect": service.inspect,
        "office_read": service.read,
        "office_find": service.find,
        "office_edit": service.edit,
        "office_create": service.create,
        "office_convert": service.convert,
    }
    if name not in calls:
        return {"status": "refused", "code": "invalid_request", "message": f"No tool named {name}"}
    schema = next(t["inputSchema"] for t in TOOLS if t["name"] == name)
    from .service import validate

    problem = validate(schema, arguments)
    if problem:
        return {"status": "refused", "code": "invalid_request", "message": problem}
    return calls[name](**arguments)


def build(config: Config) -> Server:
    service = Service(config)

    async def list_tools(ctx, params):
        return types.ListToolsResult(tools=[
            types.Tool(name=t["name"], description=t["description"], inputSchema=t["inputSchema"], outputSchema=RESULT,
                       annotations=types.ToolAnnotations(readOnlyHint=t["readOnly"], destructiveHint=False, idempotentHint=t["readOnly"]))
            for t in TOOLS])

    async def call_tool(ctx, params):
        # The libraries are synchronous; run them off the event loop so the session stays responsive.
        result = await anyio.to_thread.run_sync(dispatch, service, params.name, dict(params.arguments or {}))
        return types.CallToolResult(content=[types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False))],
                                    structuredContent=result, isError=result.get("status") == "error")

    return Server("officemcp", version=VERSION, instructions=INSTRUCTIONS, on_list_tools=list_tools, on_call_tool=call_tool)


async def serve(config: Config) -> None:
    server = build(config)
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())
