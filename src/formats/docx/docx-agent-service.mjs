import { layoutDocument, queryPage } from "./docx-layout-service.mjs";
import { executeDocx, viewOf, withTargets } from "./docx-service.mjs";
import { asView } from "../common/document-workflow.mjs";
import { docxCapabilityHelp } from "./docx-agent-help.mjs";
import {
  digest,
  newId,
  regularFile,
  sourcePath,
} from "../../infrastructure/files.mjs";
import { requireThat, OfficeError } from "../../infrastructure/errors.mjs";
import {
  DOCX_ACTION_REGISTRY,
  parseDocxActionArguments,
} from "./docx-agent-actions.mjs";

function mutationView(view, result) {
  return {
    ...result,
    view: asView(
      view.documentId,
      result.data.candidateRevision,
      result.data.candidateId,
    ),
    data: {
      ...result.data,
      nextAction:
        "Use the returned view. Save validates automatically; read or query when the task needs more evidence.",
    },
  };
}

function targetIdentity(target) {
  if (typeof target === "string") return target;
  return String(
    target?.nodeId ??
      target?.id ??
      target?.address?.nodeId ??
      target?.selector?.nodeId ??
      target?.kind ??
      "resolved-target",
  );
}

function summarizedStep(step) {
  return {
    action: step.action,
    resolvedTargets: (step.resolvedTargets ?? []).map(targetIdentity),
    verificationPassed: step.receipt?.verificationPassed !== false,
  };
}

function compactApplyResult(result, service) {
  const data = result?.data ?? {};
  const receipt = data.receipt;
  const steps = data.steps ?? receipt?.steps;
  return {
    ...result,
    data: {
      ...(data.preview !== undefined ? { preview: data.preview } : {}),
      ...(data.applied !== undefined ? { applied: data.applied } : {}),
      ...(data.wouldChange !== undefined
        ? { wouldChange: data.wouldChange }
        : {}),
      ...(data.action || receipt?.intent
        ? { action: data.action ?? receipt.intent }
        : {}),
      ...(Array.isArray(steps)
        ? { steps: steps.map(summarizedStep) }
        : {}),
      ...(Array.isArray(data.affectedTargets)
        ? { affectedTargets: data.affectedTargets.map(targetIdentity) }
        : {}),
      ...(receipt
        ? { verificationPassed: receipt.verificationPassed !== false }
        : {}),
      ...(data.validationPassed !== undefined
        ? { validationPassed: data.validationPassed }
        : {}),
      ...(data.published !== undefined ? { published: data.published } : {}),
      ...(data.created !== undefined ? { created: data.created } : {}),
      ...(data.sourceTemplateUnchanged !== undefined
        ? { sourceTemplateUnchanged: data.sourceTemplateUnchanged }
        : {}),
      ...(data.outputRevision ? { outputRevision: data.outputRevision } : {}),
      ...(data.assetId ? { assetId: data.assetId } : {}),
      ...(data.path ? { path: data.path } : {}),
      ...(data.engine ? { engine: data.engine } : {}),
      ...(data.nextAction ? { nextAction: data.nextAction } : {}),
      ...(data.atomicity ? { atomicity: data.atomicity } : {}),
      ...(data.workflow ? { workflow: data.workflow } : {}),
      ...(data.postconditions ? { postconditions: data.postconditions } : {}),
      ...(data.semanticDelta ? { semanticDelta: data.semanticDelta } : {}),
      ...(data.workflowVerification ? { workflowVerification: data.workflowVerification } : {}),
      ...(data.readback ? { readback: service && result.view && !data.preview
        ? withTargets(service, result.view, data.readback) : data.readback } : {}),
    },
  };
}

function compactValidation(result) {
  const evidence = result.evidence;
  return {
    passed: result.passed,
    revision: result.revision,
    profile: result.profile,
    checks: result.checks,
    notVerified: result.notVerified ?? [],
    ...(evidence
      ? {
          packageSummary: {
            removedParts: evidence.removedParts?.length ?? 0,
            addedParts: evidence.addedParts?.length ?? 0,
            changedParts: evidence.changedParts?.length ?? 0,
            preservedBinaryMembers: evidence.binaryMembersPreserved ?? 0,
            missingRelationshipTargets:
              evidence.relationshipGraph?.missingTargets ?? 0,
            newlyOrphanedMedia: evidence.newlyOrphanedMedia?.length ?? 0,
            unrelatedBinaryChanges:
              (evidence.binaryChanges?.unauthorizedSourceChanges?.length ?? 0) +
              (evidence.binaryChanges?.unauthorizedAddedParts?.length ?? 0),
          },
        }
      : {}),
  };
}

function compileWorkflow(request) {
  const workflow = request.workflow;
  if (!workflow) return null;
  if (["update_table_by_key", "replace_section"].includes(workflow.type)) {
    requireThat(!(request.selectorTarget && request.select), "INVALID_ARGUMENT", "Use selectorTarget or select, not both.");
    requireThat(!request.selectorTargets && !request.story, "INVALID_ARGUMENT", "These workflows accept one body target only.");
    requireThat(workflow.type !== "replace_section" || (!request.selectorTarget && !request.select), "INVALID_ARGUMENT", "replace_section resolves its own headingPath; do not supply another selector.");
    if (workflow.type === 'replace_section' && workflow.blocks?.some(b => b.kind === 'retain')) return { steps: [{ action: 'replace_section_content', arguments: { headingPath: workflow.headingPath, subsections: workflow.subsections, expectedText: workflow.expectedText, blocks: workflow.blocks } }], packageWorkflow: true };
    return { steps: [], semanticWorkflow: workflow };
  }
  if (workflow.type === "fill_template") {
    const beforeOccurrences = Object.fromEntries(
      workflow.bindings.map((binding) => [
        binding.placeholder,
        binding.expectedCount,
      ]),
    );
    return {
      steps: [
        {
          action: "replace_text",
          arguments: {
            edits: workflow.bindings.map((binding) => ({
              find: binding.placeholder,
              replace: binding.value,
            })),
            caseSensitive: true,
            changeMode: workflow.changeMode,
          },
        },
      ],
      expect: {
        beforeOccurrences,
        remainingOccurrences: Object.fromEntries(
          workflow.bindings.map((binding) => [binding.placeholder, 0]),
        ),
        changedOccurrences: workflow.bindings.reduce(
          (sum, binding) => sum + binding.expectedCount,
          0,
        ),
      },
    };
  }
  if (workflow.type === "update_table")
    return {
      steps: [
        {
          action: "write_table_cells",
          selectorTarget: request.selectorTarget,
          arguments: { cells: workflow.cells },
        },
        ...(workflow.caption
          ? [
              {
                action: "insert_caption",
                selectorTarget: request.selectorTarget,
                arguments: workflow.caption,
              },
            ]
          : []),
      ],
    };
  if (workflow.type === "normalize_document_style")
    return {
      steps: [
        ...(workflow.fontSize
          ? [
              {
                action: "normalize_body_font_size",
                arguments: {
                  fontSize: workflow.fontSize,
                  changeMode: workflow.changeMode,
                },
              },
            ]
          : []),
        ...(workflow.fontFamily
          ? [
              {
                action: "set_font_family",
                arguments: {
                  fontFamily: workflow.fontFamily,
                  changeMode: workflow.changeMode,
                },
              },
            ]
          : []),
      ],
    };
  throw new OfficeError("INVALID_ARGUMENT", "Unknown DOCX workflow.");
}

function resolveSelectorTarget(service, view, selectorTarget) {
  return service.documents.target(view, selectorTarget);
}

function workflowPostconditions(required, supplied) {
  if (!required) return supplied;
  if (!supplied) return required;
  const merged = { ...required, ...supplied };
  for (const key of ["beforeOccurrences", "remainingOccurrences"]) {
    for (const [text, count] of Object.entries(required[key] ?? {}))
      requireThat(supplied[key]?.[text] === undefined || supplied[key][text] === count,
        "INVALID_ARGUMENT", `Explicit ${key} conflicts with the workflow's required guard.`);
    merged[key] = { ...supplied[key], ...required[key] };
  }
  requireThat(supplied.changedOccurrences === undefined || required.changedOccurrences === undefined || supplied.changedOccurrences === required.changedOccurrences,
    "INVALID_ARGUMENT", "Explicit changedOccurrences conflicts with the workflow's required guard.");
  return merged;
}

function resolveImage(service, args) {
  requireThat(
    Boolean(args.assetId) !== Boolean(args.path),
    "INVALID_ARGUMENT",
    "Provide exactly one image assetId or path.",
  );
  let file;
  if (args.path) file = sourcePath(service.config, args.path);
  else {
    const asset = service.get("asset", args.assetId);
    service.doc(asset.documentId);
    requireThat(
      ["image/png", "image/jpeg"].includes(asset.mediaType),
      "UNSUPPORTED_FORMAT",
      "Image assets must be PNG or JPEG.",
    );
    requireThat(
      digest(regularFile(asset.path, 8 * 1024 * 1024)) === asset.hash,
      "INTEGRITY_FAILED",
      "Image asset changed.",
    );
    file = asset.path;
  }
  const bytes = regularFile(file, 8 * 1024 * 1024);
  const png = bytes
    .subarray(0, 8)
    .equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
  const jpeg = bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
  requireThat(png || jpeg, "UNSUPPORTED_FORMAT", "Use PNG or JPEG image bytes.");
  return {
    file,
    hash: digest(bytes),
    mediaType: png ? "image/png" : "image/jpeg",
  };
}

function prepareAgentStep(service, view, raw) {
  const entry = DOCX_ACTION_REGISTRY[raw.action];
  requireThat(entry, "INVALID_ARGUMENT", `Unknown DOCX action: ${raw.action}.`);
  let args;
  try {
    args = parseDocxActionArguments(raw.action, raw.arguments ?? {});
  } catch (error) {
    throw new OfficeError("INVALID_ARGUMENT", error.message);
  }
  requireThat(
    !(raw.selectorTarget && raw.selectorTargets),
    "INVALID_ARGUMENT",
    "Use selectorTarget or selectorTargets, not both.",
  );
  requireThat(
    !(raw.select && (raw.selectorTarget || raw.selectorTargets)),
    "INVALID_ARGUMENT",
    "Use an inline select or opaque selector targets, not both.",
  );
  const targetIds = raw.selectorTargets ??
    (raw.selectorTarget ? [raw.selectorTarget] : []);
  const targetLocators = targetIds.map((target) =>
    resolveSelectorTarget(service, view, target),
  );
  let image;
  if (["insert_image", "replace_image"].includes(raw.action)) {
    image = resolveImage(service, args);
    delete args.assetId;
    delete args.path;
  }
  return {
    action: raw.action,
    arguments: args,
    ...(raw.action === 'replace_section_content' ? { targetLocator: { kind: 'opc_object', objectKind: 'section', objectId: 'section:' + JSON.stringify(args.headingPath), partName: 'word/document.xml' } } : {}),
    ...(targetLocators[0] ? { targetLocator: targetLocators[0] } : {}),
    ...(targetLocators.length ? { targetLocators } : {}),
    ...(raw.select ? { select: raw.select } : {}),
    ...(raw.story ? { story: raw.story } : {}),
    ...(image ? { image } : {}),
  };
}

export async function executeDocxAgent(service, name, a, signal) {
  if (name === "docx_inspect" && !a.view) return { data: { ...docxCapabilityHelp(a), documentInspected: false } };
  if (name === "docx_open")
    return executeDocx(
      service,
      "docx_session",
      { action: "open", path: a.path, disposition: "retain" },
      signal,
    );
  if (name === "docx_create")
    return executeDocx(service, "docx_create", a, signal);
  if (name === "docx_save")
    return executeDocx(service, "docx_save", a, signal);
  if (name === "docx_close")
    return executeDocx(
      service,
      "docx_session",
      { action: "close", view: a.view, disposition: a.disposition },
      signal,
    );
  if (name === "docx_operation_get")
  {
    const status = await service.execute(
      "office_operation_status",
      { operationId: a.operationId },
      signal,
    );
    const stored = status.data;
    if (stored?.data?.format === "pdf" && stored.data.published) return stored;
    if (!stored?.data?.candidateId) return status;
    const candidate = service.get("candidate", stored.data.candidateId);
    return compactApplyResult({
      ...stored,
      operationId: a.operationId,
      view: asView(
        candidate.documentId,
        stored.data.candidateRevision ?? stored.data.outputRevision ?? candidate.revision,
        stored.data.candidateId,
      ),
    }, service);
  }
  if (name === "docx_apply" && a.mode === "apply") {
    const operation = service.durable.lookup(name, a);
    if (operation.replayed) {
      return compactApplyResult(mutationView(a.view, operation.result), service);
    }
  }

  if (name === "docx_layout") return layoutDocument(service, a, signal);
  if (name === "docx_query" && a.select.type === "page") {
    requireThat(!a.withinTarget && !a.story, "INVALID_ARGUMENT", "Page queries use the recorded layout; do not combine story/withinTarget.");
    return queryPage(service, a, signal);
  }
  const v = viewOf(service, a.view);
  if (name === "docx_export") {
    return service.documents.exportArtifact(name, a, "pdf", async () => {
      requireThat(service.config.engines?.libreOffice, "ENGINE_UNAVAILABLE",
        "Configure engines.libreOffice before exporting DOCX to PDF.");
      const output = service.internal("exports", newId("docx-export") + ".pdf");
      const converted = await service.worker("office_export_pdf", v.file, {},
        { format: "docx", output, bypassDocx: true }, signal);
      const bytes = regularFile(output, service.config.maxFileBytes);
      requireThat(bytes.subarray(0, 5).toString() === "%PDF-", "ENGINE_FAILED",
        "Conversion did not produce a valid PDF signature.");
      return { file: output, mediaType: "application/pdf",
        engine: converted.engine, notVerified: converted.notVerified };
    }, signal);
  }

  if (name === "docx_inspect") {
    const result = await service.worker(
      "docx_native_agent_inspect",
      v.file,
      a,
      {},
      signal,
    );
    return { view: a.view, data: { ...result, ...docxCapabilityHelp(a), documentInspected: true } };
  }
  if (name === "docx_read" && a.representation === "domain" && ['objects', 'table_cells', 'section_content'].includes(a.domain) || name === "docx_query" && a.select.type === "entity" && ['objects', 'table_cells', 'section_content'].includes(a.select.domain)) {
    requireThat(!a.section || name === 'docx_read' && a.domain === 'section_content', 'INVALID_ARGUMENT', 'section is supported only for section_content reads.');
    requireThat(!a.story && !a.withinTarget, "UNSUPPORTED_FEATURE", "Object identities use package parts; story/container scopes need explicit object qualification.");
    const result = await service.worker(name === "docx_read" ? "docx_opc_agent_read" : "docx_opc_agent_query", v.file, a, { format: "docx", bypassDocx: true }, signal);
    viewOf(service, a.view);
    return { view: a.view, data: withTargets(service, a.view, result), coverage: result.coverage };
  }
  if (name === "docx_read" || name === "docx_query") {
    requireThat(!a.section, 'INVALID_ARGUMENT', 'section is supported only with representation domain and domain section_content.');
    let workerArgs =
      name === "docx_query" && a.withinTarget
        ? {
            ...a,
            withinLocator: resolveSelectorTarget(
              service,
              a.view,
              a.withinTarget,
            ),
          }
        : a;
    let readSnapshotId = a.readSnapshotId;
    if (name === "docx_read" && a.representation === "complete") {
      const identity = {
        view: a.view,
        representation: a.representation,
        story: a.story,
        domain: a.domain,
        includeAnnotations: a.includeAnnotations,
      };
      readSnapshotId = service.documents.readSnapshot(a.view, identity, readSnapshotId);
      workerArgs = { ...workerArgs, readSnapshotId };
    }
    const result = await service.worker(
      name === "docx_read"
        ? "docx_native_agent_read"
        : "docx_native_agent_query",
      v.file,
      workerArgs,
      {},
      signal,
    );
    const data = withTargets(service, a.view, result);
    delete data.coverage;
    delete data.engineSession;
    if (readSnapshotId) data.readSnapshotId = readSnapshotId;
    return {
      view: a.view,
      data,
      coverage: result.coverage ?? {
        completeDocument: false,
        notChecked: ["visual-layout"],
      },
    };
  }
  if (name === "docx_apply") {
    requireThat(!a.workflow || a.select?.type !== "page", "INVALID_ARGUMENT", "Page selectors resolve text ranges; select a semantic table or heading for a composite workflow.");
    const compiledWorkflow = compileWorkflow(a);
    const expect = workflowPostconditions(compiledWorkflow?.expect, a.expect);
    const rawSteps =
      compiledWorkflow?.steps ??
      a.steps ?? [
        {
          action: a.action,
          arguments: a.arguments ?? {},
          ...(a.selectorTarget ? { selectorTarget: a.selectorTarget } : {}),
          ...(a.selectorTargets ? { selectorTargets: a.selectorTargets } : {}),
          ...(a.select ? { select: a.select } : {}),
          ...(a.story ? { story: a.story } : {}),
        },
      ];
    const resolvedSteps = [];
    for (const step of rawSteps) {
      if (step.select?.type === "entity" && ['objects', 'table_cells'].includes(step.select.domain)) {
        const result = await service.worker("docx_opc_agent_query", v.file, { view: a.view, select: step.select, offset: 0, limit: 2, require: "exactlyOne" }, { format: "docx", bypassDocx: true }, signal);
        const matched = withTargets(service, a.view, result).items[0];
        requireThat(matched?.target, "UNSUPPORTED_FEATURE", "The chosen object has no qualified mutation target.");
        resolvedSteps.push({ ...step, select: undefined, selectorTarget: matched.target });
      } else if (step.select?.type === "page") {
        const page = await queryPage(service, { view: a.view, select: step.select, offset: 0, limit: 2, require: "exactlyOne" }, signal);
        resolvedSteps.push({ ...step, select: undefined, selectorTarget: page.data.items[0].target });
      } else resolvedSteps.push(step);
    }
    const steps = resolvedSteps.map((step) =>
      prepareAgentStep(service, a.view, step),
    );
    if (steps.length === 1 && steps[0].action === 'replace_table_cell_text') {
      const step = steps[0];
      let available;
      try { available = await service.worker('docx_native_agent_query', v.file, { packageCellEdit: { locator: step.targetLocator, arguments: step.arguments } }, {}, signal); }
      catch (error) { if (!['ENGINE_FAILED', 'UNSUPPORTED_FEATURE'].includes(error.code)) throw error; available = { available: false }; }
      if (available.available) steps[0] = { action: 'replace_text', arguments: { edits: [{ find: step.arguments.find, replace: step.arguments.replace }], caseSensitive: true, changeMode: 'direct' }, targetLocator: available.locator, targetLocators: [available.locator] };
    }
    const objectPlan = steps.some(step => DOCX_ACTION_REGISTRY[step.action].handler === "opc");
    requireThat(!steps.some(step => step.action === 'replace_section_content') || steps.length === 1, 'INVALID_ARGUMENT', 'Section replacement is one isolated action; edit retained objects in a subsequent revision.');
    requireThat(!objectPlan || steps.every(step => DOCX_ACTION_REGISTRY[step.action].handler === "opc") && (!a.workflow || compiledWorkflow?.packageWorkflow) && !expect,
      "INVALID_ARGUMENT", "Use an object-only batch for guarded OOXML actions; SDK actions, workflows and generic text postconditions need a separate phase.");
    const nativeArgs = {
      documentId: a.view.documentId,
      expectedSourceRevision: v.d.revision,
      preservationProfile: objectPlan ? "docx.guarded-objects.v1" : "docx.superdoc-agent.v1",
      mode: a.mode,
      steps,
      author: a.author,
      ...(a.readback ? { readback: a.readback } : {}),
      ...(compiledWorkflow?.semanticWorkflow ? {
        semanticWorkflow: compiledWorkflow.semanticWorkflow,
        workflowLocator: a.selectorTarget ? resolveSelectorTarget(service, a.view, a.selectorTarget) : undefined,
        workflowSelect: a.select,
      } : {}),
      ...(a.workflow ? { workflow: a.workflow.type } : {}),
      ...(expect ? { expect } : {}),
      ...(a.view.candidateId
        ? {
            candidateId: a.view.candidateId,
            expectedCandidateRevision: a.view.revision,
          }
        : {}),
    };
    if (a.mode === "preview") {
      const preview = await service.worker(
        objectPlan ? "docx_opc_agent_preview" : "docx_native_agent_preview",
        v.file,
        nativeArgs,
        objectPlan ? { format: "docx", bypassDocx: true } : {},
        signal,
      );
      return compactApplyResult({ view: a.view, data: preview });
    }
    nativeArgs.operationId = a.operationId;
    return compactApplyResult(
      mutationView(
        a.view,
        await service.operation(objectPlan ? "docx_opc_agent_apply" : "docx_native_agent_apply", nativeArgs, signal, {
          transition: true,
          ...(objectPlan ? { workerContext: { format: "docx", bypassDocx: true } } : {}),
          request: { name, args: a },
        }),
      ), service,
    );
  }
  if (name === "docx_validate") {
    requireThat(
      v.c,
      "INVALID_ARGUMENT",
      "docx_validate requires an unpublished candidate view.",
    );
    const validation = await service.documents.validate({
      documentId: a.view.documentId,
      candidateId: a.view.candidateId,
      expectedCandidateRevision: a.view.revision,
    }, signal);
    return {
      view: a.view,
      data: {
        validationId: validation.id,
        ...compactValidation(validation.result),
      },
    };
  }

  if (name === "docx_compare") {
    const result = await service.worker(
      "docx_native_agent_compare",
      v.file,
      a,
      { original: v.d.snapshot },
      signal,
    );
    return { view: a.view, data: result };
  }
  if (name === "docx_image_get") {
    const locator = resolveSelectorTarget(service, a.view, a.target);
    requireThat(
      locator?.kind === "image" && typeof locator.partName === "string",
      "INVALID_ARGUMENT",
      "Use an image target returned by docx_read(domain=images).",
    );
    const artifact = await service.artifacts.produce({ kind: "office_extract", input: v.file,
      args: { source: { kind: "package_part", name: locator.partName } },
      context: { format: "docx", bypassDocx: true }, documentId: a.view.documentId,
      prefix: "docx-image", reservation: 8 * 1024 * 1024, requireImage: true, includeImage: true }, signal);
    return { ...artifact, view: a.view,
      data: { ...artifact.data, engine: "exact OOXML package extraction" } };
  }
  requireThat(false, "UNKNOWN_TOOL", "Unknown DOCX agent tool.");
}
