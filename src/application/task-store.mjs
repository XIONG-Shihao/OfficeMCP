import { operationDescriptor } from "../contracts/registry.mjs";
import { fingerprint, newId } from "../infrastructure/files.mjs";
import { requireThat } from "../infrastructure/errors.mjs";

/** Modern tasks have their own durable record kind. Legacy jobs are an adapter. */
export class TaskStore {
  constructor(service) {
    this.service = service;
    this.active = new Map();
  }

  get(taskId) {
    return this.service.get("task", taskId);
  }

  cancel(taskId) {
    const record = this.get(taskId);
    this.active.get(taskId)?.abort();
    return {
      accepted: true,
      state: this.active.has(taskId) ? "cancelling" : record.status,
    };
  }

  start(name, args, { legacy = false } = {}) {
    const kind = legacy ? "job" : "task";
    const taskId = newId(legacy ? "job" : "task");
    const controller = new AbortController();
    const service = this.service;
    service.checkQuota();
    requireThat(
      this.active.size < 2,
      "BUSY",
      "Background task limit reached.",
    );
    this.active.set(taskId, controller);
    service.put(kind, taskId, {
      status: "running",
      method: "tools/call",
      operation: name,
      argumentsFingerprint: fingerprint(args),
      resultContract:
        operationDescriptor(service.config.toolset, name)?.policy.resultContract ?? "tool-result",
      cancellation: "available",
      progress: { completed: 0, total: 1 },
      createdAt: new Date(service.now()).toISOString(),
    });
    void service
      .execute(name, { ...args, background: false }, controller.signal)
      .then(({ nativeImage, ...data }) => {
        if (service.closed) return;
        const result = service.envelope(data, name);
        requireThat(
          Buffer.byteLength(JSON.stringify(result)) * 2 <
            service.config.maxResponseBytes,
          "LIMIT_EXCEEDED",
          "Task result exceeds response budget; request a smaller window.",
        );
        service.put(kind, taskId, {
          ...service.get(kind, taskId),
          status: "completed",
          cancellation: "closed",
          progress: { completed: 1, total: 1 },
          result,
          content: nativeImage ? [nativeImage] : [],
        });
      })
      .catch((error) => {
        if (!service.closed)
          service.put(kind, taskId, {
            ...service.get(kind, taskId),
            status: error.code === "CANCELLED" ? "cancelled" : "failed",
            cancellation: "closed",
            result: service.failure(error, args, name).payload,
          });
      })
      .finally(() => this.active.delete(taskId));
    return taskId;
  }
}
