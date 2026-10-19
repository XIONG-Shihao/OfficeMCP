import { OfficeError, requireThat, errorPayload } from "../infrastructure/errors.mjs";
import { fingerprint } from "../infrastructure/files.mjs";

/** The single state machine for idempotent writes and crash-safe publication. */
export class DurableOperations {
  constructor(service) {
    this.service = service;
  }

  lookup(name, args, options = {}) {
    const id = args.operationId;
    const requestFingerprint = fingerprint({
      name: options.fingerprintName ?? name,
      args: options.fingerprintArgs ?? args,
    });
    const existing = this.service.get("operation", id, false);
    if (existing) {
      requireThat(
        existing.fingerprint === requestFingerprint,
        "OPERATION_ID_CONFLICT",
        "Operation ID was already used with different arguments.",
      );
      if (existing.status === "expired")
        throw new OfficeError(
          "OPERATION_EXPIRED",
          "Receipt expired; this operation ID cannot be reused.",
        );
      if (existing.status === "failed")
        throw new OfficeError(existing.error.code, existing.error.message);
      if (existing.status === "completed")
        return {
          id,
          fingerprint: requestFingerprint,
          replayed: true,
          result: existing.result,
        };
      throw new OfficeError(
        "OPERATION_PENDING",
        "Operation is running or requires recovery; inspect the existing operation ID.",
      );
    }
    return { id, fingerprint: requestFingerprint, replayed: false };
  }

  begin(name, args, options = {}) {
    const operation = this.lookup(name, args, options);
    if (!operation.replayed) this.write({ ...operation, name }, "running");
    return operation;
  }

  write(operation, status, extra = {}) {
    this.service.put("operation", operation.id, {
      fingerprint: operation.fingerprint,
      name: operation.name,
      status,
      ...extra,
    });
  }

  publishing(operation, publication) {
    this.write(operation, "publishing", { publication });
  }

  complete(operation, result, publication) {
    this.write(operation, "completed", {
      result,
      ...(publication ? { publication } : {}),
    });
    operation.completed = true;
    return result;
  }

  fail(operation, error) {
    this.write(operation, "failed", { error: errorPayload(error) });
  }

  async run(name, args, work, options = {}) {
    const operation = { ...this.begin(name, args, options), name };
    if (operation.replayed) return operation.result;
    const context = {
      operation,
      publishing: (publication) => this.publishing(operation, publication),
      complete: (result, publication) =>
        this.complete(operation, result, publication),
    };
    try {
      const result = await work(context);
      return operation.completed ? result : this.complete(operation, result);
    } catch (error) {
      if (this.service.db.isTransaction) this.service.db.exec("ROLLBACK");
      const current = this.service.get("operation", operation.id);
      if (current.status === "completed") return current.result;
      if (current.status === "publishing") {
        this.service.reconcile(operation.id, current);
        const recovered = this.service.get("operation", operation.id);
        if (recovered.status === "completed") return recovered.result;
        throw new OfficeError(
          "PUBLICATION_UNCERTAIN",
          "Publication needs reconciliation; inspect the existing operation ID.",
        );
      }
      this.fail(operation, error);
      throw error;
    }
  }
}
