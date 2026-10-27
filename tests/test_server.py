"""The real MCP server over stdio, driven by the MCP SDK's own client."""

import json
import sys
import unittest

import anyio
from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

from .support import Case, make_docx


class ServerTests(Case):
    def test_a_client_lists_the_seven_tools_and_edits_a_document(self):
        path = make_docx(self.inputs)
        config = self.root / "config.json"
        config.write_text(json.dumps({"roots": [str(self.inputs)]}))

        async def session():
            parameters = StdioServerParameters(command=sys.executable, args=["-m", "officemcp", "serve", "--config", str(config)])
            async with stdio_client(parameters) as (read, write):
                async with ClientSession(read, write) as client:
                    await client.initialize()
                    tools = (await client.list_tools()).tools
                    inspected = (await client.call_tool("office_inspect", {"path": str(path)})).structured_content
                    edited = (await client.call_tool("office_edit", {"path": str(path), "expectedRevision": inspected["revision"],
                                                                     "actions": [{"action": "set_properties", "title": "From MCP"}]})).structured_content
                    refused = (await client.call_tool("office_read", {"path": "/etc/hosts"})).structured_content
                    bad = (await client.call_tool("office_find", {"path": str(path)})).structured_content
                    return tools, inspected, edited, refused, bad

        tools, inspected, edited, refused, bad = anyio.run(session)
        self.assertEqual(len(tools), 7)
        self.assertTrue(next(t for t in tools if t.name == "office_inspect").annotations.read_only_hint)
        self.assertEqual(inspected["status"], "ok")
        self.assertEqual(edited["status"], "ok", edited)
        self.assertEqual(edited["path"], str(path), "saved in place")
        self.assertEqual(refused["code"], "path_not_allowed")
        self.assertEqual(bad["code"], "invalid_request")


if __name__ == "__main__":
    unittest.main()
