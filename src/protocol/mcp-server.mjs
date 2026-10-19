import {
  ProtocolError,
  ProtocolErrorCode,
  Server,
} from "@modelcontextprotocol/server";
import { manifest } from "../contracts/registry.mjs";
import { MCP_SERVER_INFO, TASKS_EXTENSION } from "./mcp-protocol.mjs";

function modernManifest(tools) {
  return tools
    .filter(
      (tool) =>
        !["office_job_status", "office_job_cancel"].includes(tool.name),
    )
    .map((tool) => {
      const copy = structuredClone(tool);
      if (copy.inputSchema.properties)
        delete copy.inputSchema.properties.background;
      if (Array.isArray(copy.inputSchema.required))
        copy.inputSchema.required = copy.inputSchema.required.filter(
          (name) => name !== "background",
        );
      return copy;
    });
}

export function createMcpServer(service, factoryContext = { era: "legacy" }) {
  const all = service.manifest?.() ?? manifest();
  const exposed = factoryContext.era === "modern" ? modernManifest(all) : all;
  const server = new Server(
    MCP_SERVER_INFO,
    {
      capabilities: {
        tools: {},
        ...(factoryContext.era === "modern"
          ? { extensions: { [TASKS_EXTENSION]: {} } }
          : {}),
      },
      cacheHints: {
        "server/discover": { ttlMs: 60_000, cacheScope: "private" },
        "tools/list": { ttlMs: 60_000, cacheScope: "private" },
      },
    },
  );
  server.setRequestHandler("tools/list", async () => ({
    tools: exposed,
  }));
  server.setRequestHandler("tools/call", async (request, context) => {
    if (!exposed.some((tool) => tool.name === request.params.name))
      throw new ProtocolError(
        ProtocolErrorCode.InvalidParams,
        "Unknown tool",
      );
    const result = await service.call(
      request.params.name,
      request.params.arguments ?? {},
      context.mcpReq.signal,
    );
    if (
      factoryContext.era === "modern" &&
      request.params.name === "office_server_info" &&
      result.payload?.data
    )
      result.payload.data.operations = exposed.map((tool) => tool.name);
    return {
      isError: result.isError ?? false,
      structuredContent: result.payload,
      content: [
        { type: "text", text: JSON.stringify(result.payload) },
        ...(result.content ?? []),
      ],
    };
  });
  return server;
}
