import { z } from "zod";

export function defineOperation(spec) {
  const {
    name, description, input, output, risk, handler, execute, contract, presets,
    execution = "synchronous",
    engines = ["format-dependent"],
    destructive = false,
    discovery = {},
  } = spec;
  if (!name || !description || !input || !output || !handler || typeof execute !== "function")
    throw new TypeError(`Incomplete operation specification: ${name ?? "unnamed"}`);
  return Object.freeze({
    name,
    description,
    input,
    output,
    execute,
    discovery: Object.freeze({
      resultType: "JSON",
      wireTypes: ["structuredContent: application/json", "content: text containing the same JSON"],
      status: "Implemented",
      ...discovery,
    }),
    policy: Object.freeze({ risk, execution, handler, supportedEngines: engines,
      resultContract: contract, presets, destructive }),
  });
}

export function parseOperationArguments(spec, raw) {
  return (spec.input.strict?.() ?? spec.input).safeParse(raw);
}

export function operationManifestEntry(spec) {
  const { destructive, ...metadata } = spec.policy;
  return {
    name: spec.name,
    description: spec.description,
    inputSchema: {
      type: "object",
      ...z.toJSONSchema(spec.input, { target: "draft-2020-12" }),
    },
    outputSchema: spec.output,
    annotations: {
      readOnlyHint: spec.policy.risk === "read",
      destructiveHint: spec.policy.destructive,
      openWorldHint: false,
    },
    _meta: { "io.officemcp/operation": { name: spec.name, ...metadata } },
  };
}
