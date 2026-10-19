import { operationDescriptor } from "../contracts/registry.mjs";

export function executionPolicy(service, name) {
  const operation = operationDescriptor(service.config.toolset, name);
  return operation?.policy.execution === "taskEligible"
    ? "task-eligible"
    : "synchronous";
}
