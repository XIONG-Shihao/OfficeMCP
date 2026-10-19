import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import {
  MCP_PROTOCOL_VERSION,
  MCP_SERVER_INFO,
  TASKS_EXTENSION,
} from "./mcp-protocol.mjs";
import { executionPolicy } from "../application/operation-policy.mjs";

const INVALID_PARAMS = -32602;
const INTERNAL_ERROR = -32603;
const MISSING_REQUIRED_CLIENT_CAPABILITY = -32021;
const POLL_INTERVAL_MS = 500;

const isObject = (value) =>
  value !== null && typeof value === "object" && !Array.isArray(value);

function isTaskEligible(service, name) {
  return executionPolicy(service, name) === "task-eligible";
}

function requestMeta(message) {
  return isObject(message.params) && isObject(message.params._meta)
    ? message.params._meta
    : null;
}

function isModern(message) {
  return (
    requestMeta(message)?.["io.modelcontextprotocol/protocolVersion"] ===
    MCP_PROTOCOL_VERSION
  );
}

function hasTasksCapability(message) {
  return Boolean(
    requestMeta(message)?.["io.modelcontextprotocol/clientCapabilities"]
      ?.extensions?.[TASKS_EXTENSION],
  );
}

function serverMeta() {
  return {
    "io.modelcontextprotocol/serverInfo": MCP_SERVER_INFO,
  };
}

function toolResult(payload, isError = false, content = []) {
  return {
    resultType: "complete",
    isError,
    structuredContent: payload,
    content: [{ type: "text", text: JSON.stringify(payload) }, ...content],
  };
}

function failedPayload(service, job) {
  if (job.result?.contractVersion) return job.result;
  return service.envelope({
    status: "failed",
    error: job.result?.error ?? {
      code: "WORKER_FAILED",
      message: "Background work failed.",
    },
  }, job.operation);
}

function taskFromRecord(service, taskId, job) {
  const base = {
    taskId,
    createdAt:
      job.createdAt ?? new Date(job.updatedAt ?? service.now()).toISOString(),
    lastUpdatedAt: new Date(job.updatedAt ?? service.now()).toISOString(),
    ttlMs: service.config.artifactRetentionMs,
    pollIntervalMs: POLL_INTERVAL_MS,
  };
  if (job.status === "running") return { ...base, status: "working" };
  if (job.status === "cancelled") return { ...base, status: "cancelled" };
  if (job.status === "completed")
    return {
      ...base,
      status: "completed",
      result: toolResult(job.result, false, job.content ?? []),
    };
  return {
    ...base,
    status: "completed",
    statusMessage: "The document operation returned an error.",
    result: toolResult(failedPayload(service, job), true),
  };
}

/**
 * Adds the official io.modelcontextprotocol/tasks extension to the v2 stdio
 * entry. The upstream TypeScript SDK 2.0 core intentionally reserves but does
 * not yet dispatch the extension methods, so this narrow transport adapter
 * handles only task creation/poll/update/cancel and delegates every core MCP
 * message to serveStdio.
 */
export class TaskAwareStdioTransport {
  constructor(service, input, output) {
    this.service = service;
    this.inner = new StdioServerTransport(input, output);
  }

  async start() {
    this.inner.onclose = () => this.onclose?.();
    this.inner.onerror = (error) => this.onerror?.(error);
    this.inner.onmessage = (message, extra) => {
      void this.route(message, extra).catch((error) => {
        this.onerror?.(
          error instanceof Error ? error : new Error(String(error)),
        );
        if (message?.id !== undefined)
          void this.sendError(
            message.id,
            INTERNAL_ERROR,
            "Internal server error",
          );
      });
    };
    await this.inner.start();
  }

  async close() {
    await this.inner.close();
  }

  async send(message, options) {
    await this.inner.send(message, options);
  }

  async sendResult(id, result) {
    await this.inner.send({ jsonrpc: "2.0", id, result });
  }

  async sendError(id, code, message, data) {
    await this.inner.send({
      jsonrpc: "2.0",
      id,
      error: { code, message, ...(data === undefined ? {} : { data }) },
    });
  }

  async requireCapability(message) {
    const meta = requestMeta(message);
    const capabilities =
      meta?.["io.modelcontextprotocol/clientCapabilities"];
    const clientInfo = meta?.["io.modelcontextprotocol/clientInfo"];
    if (
      !isObject(capabilities) ||
      (clientInfo !== undefined &&
        (!isObject(clientInfo) ||
          typeof clientInfo.name !== "string" ||
          typeof clientInfo.version !== "string"))
    ) {
      await this.sendError(
        message.id,
        INVALID_PARAMS,
        "Invalid _meta envelope",
      );
      return false;
    }
    if (!hasTasksCapability(message)) {
      await this.sendError(
        message.id,
        MISSING_REQUIRED_CLIENT_CAPABILITY,
        "Missing required client capability",
        { requiredCapabilities: { extensions: { [TASKS_EXTENSION]: {} } } },
      );
      return false;
    }
    return true;
  }

  async route(message, extra) {
    if (!isModern(message) || message?.id === undefined) {
      this.onmessage?.(message, extra);
      return;
    }
    const method = message.method;
    const isToolCall = method === "tools/call";
    const taskEligible =
      isToolCall && isTaskEligible(this.service, message.params?.name);
    const useTask = taskEligible && hasTasksCapability(message);
    if (
      !useTask &&
      !["tasks/get", "tasks/update", "tasks/cancel"].includes(method)
    ) {
      if (
        isToolCall &&
        isObject(message.params?.arguments) &&
        "background" in message.params.arguments
      ) {
        const args = { ...message.params.arguments };
        delete args.background;
        this.onmessage?.(
          {
            ...message,
            params: { ...message.params, arguments: args },
          },
          extra,
        );
      } else this.onmessage?.(message, extra);
      return;
    }
    if (!(await this.requireCapability(message))) return;

    if (useTask) {
      const args = isObject(message.params.arguments)
        ? { ...message.params.arguments }
        : {};
      delete args.background;
      let taskId;
      try {
        taskId = this.service.startTask(message.params.name, args);
      } catch (error) {
        const result = this.service.failure(error, args);
        await this.sendResult(message.id, {
          resultType: "complete",
          ...toolResult(result.payload, true),
          _meta: serverMeta(),
        });
        return;
      }
      const task = taskFromRecord(
        this.service,
        taskId,
        this.service.taskStore.get(taskId),
      );
      await this.sendResult(message.id, {
        resultType: "task",
        ...task,
        _meta: serverMeta(),
      });
      return;
    }

    const taskId = message.params?.taskId;
    if (typeof taskId !== "string" || taskId.length === 0) {
      await this.sendError(message.id, INVALID_PARAMS, "Invalid taskId");
      return;
    }
    let job;
    try {
      job = this.service.taskStore.get(taskId);
    } catch {
      await this.sendError(message.id, INVALID_PARAMS, "Unknown taskId");
      return;
    }

    if (method === "tasks/get") {
      await this.sendResult(message.id, {
        resultType: "complete",
        ...taskFromRecord(this.service, taskId, job),
        _meta: serverMeta(),
      });
      return;
    }
    if (method === "tasks/cancel") {
      this.service.taskStore.cancel(taskId);
      await this.sendResult(message.id, {
        resultType: "complete",
        _meta: serverMeta(),
      });
      return;
    }
    await this.sendResult(message.id, {
      resultType: "complete",
      _meta: serverMeta(),
    });
  }
}
