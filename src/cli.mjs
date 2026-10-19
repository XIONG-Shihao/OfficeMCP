#!/usr/bin/env node
import { assertSupportedRuntime, javascriptRuntime } from "./infrastructure/node-version.mjs";

try {
  assertSupportedRuntime();
} catch (error) {
  process.stderr.write(`${error.message}\n`);
  process.exit(1);
}

const [
  { createMcpServer },
  { serveStdio },
  { OfficeService },
  { configFrom },
  { VERSION },
  { TaskAwareStdioTransport },
  { MCP_PROTOCOL_VERSION, TASKS_EXTENSION },
] = await Promise.all([
  import("./protocol/mcp-server.mjs"),
  import("@modelcontextprotocol/server/stdio"),
  import("./application/service.mjs"),
  import("./infrastructure/files.mjs"),
  import("./contracts/registry.mjs"),
  import("./protocol/tasks-extension.mjs"),
  import("./protocol/mcp-protocol.mjs"),
]);
if (process.argv[2] === "doctor") {
  console.log(
    JSON.stringify(
      {
        version: VERSION,
        javascriptRuntime: javascriptRuntime(),
        platform: `${process.platform}-${process.arch}`,
        transport: "stdio",
        protocolVersion: MCP_PROTOCOL_VERSION,
        extensions: [TASKS_EXTENSION],
        formats: ["docx", "xlsx", "csv", "pptx", "pdf"],
        sqlite: (await import("./infrastructure/sqlite.mjs")).sqliteEngine,
        securitySandbox: false,
        publicRelease: false,
      },
      null,
      2,
    ),
  );
} else {
  const index = process.argv.indexOf("--config");
  if (process.argv[2] !== "serve" || index < 0 || !process.argv[index + 1]) {
    process.stderr.write(
      "Usage: officemcp-preview serve --config /absolute/config.json | doctor\n",
    );
    process.exit(1);
  }
  let service;
  try {
    const config = configFrom(process.argv[index + 1]);
    service = new OfficeService(config);
  } catch (e) {
    process.stderr.write(
      `OfficeMCP startup failed (${e.code ?? "INVALID_CONFIG"}). Check configuration and state ownership.\n`,
    );
    process.exit(1);
  }
  const transport = new TaskAwareStdioTransport(service);
  const server = serveStdio((context) => createMcpServer(service, context), {
    legacy: "serve",
    transport,
    onerror: (error) =>
      process.stderr.write(`OfficeMCP protocol error: ${error.message}\n`),
  });
  let stopped = false;
  const close = async () => {
    if (stopped) return;
    stopped = true;
    service.close();
    await server.close();
    process.exit(0);
  };
  process.once("SIGINT", close);
  process.once("SIGTERM", close);
  process.stdin.once("end", close);
}
