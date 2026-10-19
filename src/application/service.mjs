import { executeTool, bindCommands } from "./tool-dispatch.mjs";
import { runWorker } from "../workers/worker-runner.mjs";
import { CoordinatorState } from "./coordinator-state.mjs";
import { DocumentRuntime } from "./document-runtime.mjs";
import { ArtifactRuntime } from "./artifact-runtime.mjs";
import { DurableOperations } from "./durable-operation.mjs";
import { boundWorkbookMutationEvidence } from "./workbook-mutation-evidence.mjs";
import {
  DOCX_AGENT_CONTRACT,
} from "../formats/docx/docx-agent-contract.mjs";
import * as fs from "node:fs";
import path from "node:path";
import { writeDiagnostic } from "../infrastructure/diagnostics.mjs";
import {publishArtifactManifest} from '../infrastructure/artifact-manifest.mjs';
import {
  manifest,
  operationDescriptor,
  CONTRACT,
} from "../contracts/registry.mjs";
import { parseOperationArguments } from "../contracts/operation-spec.mjs";
import { OfficeError, requireThat, errorPayload } from "../infrastructure/errors.mjs";
import {DocxSnapshotRevisions} from "../infrastructure/docx-snapshot-revisions.mjs";
import {boundSerializedValidation} from '../formats/docx/docx-serialized-heading-proof.mjs';
import {freshRevisionProof} from '../infrastructure/streaming-file-revision.mjs';
import {
  fileRevision,
  newId,
  sourcePath,
  syncDirectory,
  confined,
} from "../infrastructure/files.mjs";
export class OfficeService extends CoordinatorState {
  constructor(config, testing = {}, host = {}) {
    super(config, testing);
    this.docxSnapshotRevisions=new DocxSnapshotRevisions();
    this.commands = bindCommands(this);
    this.durable = new DurableOperations(this);
    this.documents = new DocumentRuntime(this);
    this.artifacts = new ArtifactRuntime(this);
    this.publication = host.publication ?? null;
    this.recover();
  }
  manifest() {
    return manifest(this.config.toolset);
  }
  doc(id) {
    const d = this.get("document", id);
    requireThat(!d.closed, "HANDLE_CLOSED", "Document handle is closed.");
    requireThat(
      this.now() - d.updatedAt < this.config.artifactRetentionMs,
      "HANDLE_EXPIRED",
      "Document lease expired; open a new handle.",
    );
    requireThat(
      this.config.readRoots.some((r) => confined(r, d.source)) ||
        (d.generated &&
          d.source === d.snapshot &&
          confined(
            path.join(this.config.stateDirectory, "snapshots"),
            d.snapshot,
          )),
      "ACCESS_DENIED",
      "Document is outside configured roots.",
    );
    this.put("document", id, d);
    return d;
  }
  view(args) {
    const d = this.doc(args.documentId);
    const c = args.candidateId ? this.get("candidate", args.candidateId) : null;
    requireThat(
      !c || (c.documentId === args.documentId && !c.discarded),
      "ACCESS_DENIED",
      "Candidate does not belong to this open document.",
    );
    const revision = c?.revision ?? d.revision;
    requireThat(
      args.expectedRevision === revision,
      "REVISION_CONFLICT",
      "Read requires the exact source/candidate revision.",
    );
    const file = c?.path ?? d.snapshot;
    requireThat(
      (d.format==='docx'?this.docxSnapshotRevisions.revision(file,this.config.maxFileBytes):fileRevision(file,this.config.maxFileBytes)) === revision,
      "INTEGRITY_FAILED",
      "Snapshot bytes changed.",
    );
    return { d, c, file, revision };
  }
  checkSource(d, expected) {
    const file=d.generated?d.snapshot:sourcePath(this.config,d.source);
    const proof=d.format==='docx'?freshRevisionProof(file,this.config.maxFileBytes):null;
    requireThat(
      expected === d.revision &&
        (proof?.revision??fileRevision(file,this.config.maxFileBytes)) === expected,
      "SOURCE_REVISION_CONFLICT",
      "Source changed since open; reopen and resolve the target.",
    );
    return proof?{file,identity:proof.identity}:undefined;
  }
  async worker(kind, input, args = {}, other = {}, signal) {
    return runWorker(this, kind, input, args, other, signal);
  }
  envelope(data, name) {
    return {
      contractVersion:
        (operationDescriptor(this.config.toolset, name)?.policy.resultContract === DOCX_AGENT_CONTRACT || this.config.toolset === "docx-agent")
          ? DOCX_AGENT_CONTRACT
          : CONTRACT,
      requestId: newId("req"),
      status: "completed",
      warnings: [],
      ...data,
    };
  }
  failure(e, args = {}, name) {
    const error = errorPayload(e);
    if (!(e instanceof OfficeError)) {
      const diagnosticId = writeDiagnostic(this.config, "tool-call", e);
      error.diagnosticId = diagnosticId;
      error.message = `Operation failed. Diagnostic ID: ${diagnosticId}.`;
    }
    if (this.config.toolset === "docx-agent" || operationDescriptor(this.config.toolset, name)?.policy.resultContract === DOCX_AGENT_CONTRACT) {
      if (["TARGET_CONFLICT", "REVISION_CONFLICT", "CANDIDATE_REVISION_CONFLICT"].includes(error.code))
        error.recovery =
          "Call docx_read or docx_query with the current view, then submit a new operationId with the returned target.";
      if (error.code === "OPERATION_ID_CONFLICT")
        error.recovery =
          "Call docx_operation_get for the existing receipt. An operationId is immutable; after resolving a known failed attempt, submit corrected arguments under a new operationId.";
      if (["OPERATION_PENDING", "PARTIAL_MUTATION"].includes(error.code))
        error.recovery =
          "Call docx_operation_get with the same operationId before retrying.";
      if (error.code === "ENGINE_UNAVAILABLE")
        error.recovery =
          "Ask the operator to configure the required local engine. DOCX mutations remain available through SuperDoc; PDF export requires LibreOffice.";
    }
    return {
      isError: true,
      payload: this.envelope({
        status: "failed",
        ...(args.operationId ? { operationId: args.operationId } : {}),
        error,
      }, name),
    };
  }
  parseArguments(name, raw) {
    const spec = operationDescriptor(this.config.toolset, name);
    requireThat(spec, "UNKNOWN_TOOL", "Unknown tool.");
    const parsed = parseOperationArguments(spec, raw);
    if (!parsed.success) {
      const error = new OfficeError(
        "INVALID_ARGUMENT",
        "Arguments do not match the published tool schema.",
      );
      error.argumentIssues = parsed.error.issues.map((issue) => ({
        path: issue.path.join("."),
        code: issue.code,
      }));
      throw error;
    }
    return parsed.data;
  }
  startBackground(name, args) {
    return this.taskStore.start(name, args, { legacy: true });
  }
  startTask(name, raw = {}) {
    return this.taskStore.start(name, this.parseArguments(name, raw));
  }
  async call(name, raw = {}, signal) {
    this.activeRequests++;
    try {
      const args = this.parseArguments(name, raw);
      if (args.background) {
        const id = this.startBackground(name, args);
        return {
          payload: this.envelope({
            status: "running",
            jobId: id,
            pollAfterMs: 500,
          }),
        };
      }
      const { nativeImage, ...data } = await this.execute(name, args, signal);
      const payload = this.envelope(data, name);
      requireThat(
        Buffer.byteLength(JSON.stringify(payload)) * 2 <
          this.config.maxResponseBytes,
        "LIMIT_EXCEEDED",
        "Result exceeds the response budget; request a smaller window.",
      );
      return { payload, ...(nativeImage ? { content: [nativeImage] } : {}) };
    } catch (e) {
      return this.failure(e, raw, name);
    } finally {
      this.activeRequests--;
    }
  }
  async execute(name, a, signal) {
    return executeTool(this, name, a, signal);
  }
  stateInfo() {
    const counts = Object.fromEntries(
      this.db
        .prepare("SELECT kind,COUNT(*) AS count FROM records GROUP BY kind")
        .all()
        .map((r) => [r.kind, r.count]),
    );
    return {
      schemaVersion: 3,
      counts,
      artifactRetentionMs: this.config.artifactRetentionMs,
      operationRetentionMs: this.config.operationRetentionMs,
      cleanup:
        "explicit; idle leases expire; published outputs and sources are retained",
      operationIds: "permanent tombstones; never reused",
      limits: this.configLimits(),
    };
  }
  cleanup() {
    // Refuse before touching anything: a cleanup that answers BUSY must not
    // have cancelled the work that made it busy.
    requireThat(
      ![...this.docxPool.workers].some((slot) => slot.busy),
      "BUSY",
      "DOCX host is busy.",
    );
    requireThat(
      this.activeRequests <= 1 &&
        this.activeWorkers.size === 0 &&
        this.jobs.size === 0,
      "BUSY",
      "Cleanup requires an idle coordinator.",
    );
    this.docxPool.reset();
    const expired = new Set(
      this.rows("document")
        .filter(
          (d) => this.now() - d.updatedAt >= this.config.artifactRetentionMs,
        )
        .map((d) => d.id),
    );
    let removedRecords = 0,
      expiredReceipts = 0,
      removedFiles = 0;
    this.db.exec("BEGIN IMMEDIATE");
    try {
      for (const kind of [
        "document",
        "docxTarget",
        "candidate",
        "asset",
        "validation",
        "docxReadSnapshot",
        "docxLayout",
        "cursor",
        "job",
        "task",
      ])
        for (const r of this.rows(kind)) {
          const isExpired =
            kind === "document"
              ? expired.has(r.id)
              : r.documentId
                ? expired.has(r.documentId)
                : this.now() - r.updatedAt >= this.config.artifactRetentionMs;
          if (isExpired) {
            this.delete(kind, r.id);
            removedRecords++;
          }
        }
      for (const op of this.rows("operation"))
        if (
          ["completed", "failed"].includes(op.status) &&
          this.now() - op.updatedAt >= this.config.operationRetentionMs
        ) {
          this.delete("docxReceipt", op.id);
          this.put("operation", op.id, {
            status: "expired",
            fingerprint: op.fingerprint,
            name: op.name,
          });
          expiredReceipts++;
        }
      this.db.exec("COMMIT");
    } catch (e) {
      this.db.exec("ROLLBACK");
      throw e;
    }
    const referenced = new Set([
      ...this.rows("document").map((d) => d.snapshot),
      ...this.rows("candidate")
        .filter((c) => !c.discarded)
        .flatMap((c) => [
          c.path,
          c.validationBase,
          ...(c.stages ?? []).flatMap((s) => [s.base, s.path]),
        ]),
      ...this.rows("asset").map((a) => a.path),
      ...this.rows("operation")
        .map((operation) => operation.publication?.tmp)
        .filter(Boolean),
    ]);
    // Only private artifact directories are eligible, never output roots.
    for (const dir of ["snapshots", "candidates", "assets", "exports"])
      for (const entry of fs.readdirSync(this.internal(dir), {
        withFileTypes: true,
      })) {
        const file = this.internal(dir, entry.name);
        if (entry.isFile() && !referenced.has(file)) {
          fs.unlinkSync(file);
          removedFiles++;
        }
      }
    return {
      removedRecords,
      expiredReceipts,
      removedFiles,
      expiredDocuments: expired.size,
      sourcesDeleted: 0,
      publishedOutputsDeleted: 0,
    };
  }
  configLimits() {
    const { readRoots, writeRoots, stateDirectory, engines, ...limits } =
      this.config;
    return limits;
  }
  viewIdentity(a, v) {
    return {
      id: a.documentId,
      sourceRevision: v.d.revision,
      view: v.c ? "candidate" : "source",
      candidateId: a.candidateId,
      revision: v.revision,
      format: v.d.format,
    };
  }
  candidate(a) {
    const d = this.doc(a.documentId),
      c = this.get("candidate", a.candidateId);
    requireThat(
      c.documentId === a.documentId && !c.discarded,
      "ACCESS_DENIED",
      "Candidate is unavailable for this document.",
    );
    requireThat(
      c.revision === a.expectedCandidateRevision &&
        (d.format==='docx'?this.docxSnapshotRevisions.revision(c.path,this.config.maxFileBytes):fileRevision(c.path,this.config.maxFileBytes)) === c.revision,
      "CANDIDATE_REVISION_CONFLICT",
      "Candidate revision changed.",
    );
    return { d, c };
  }
  asset(file, documentId, mediaType, verifiedRevision) {
    const id = newId("asset"),
      hash = verifiedRevision ?? fileRevision(file, this.config.maxFileBytes);
    this.put("asset", id, {
      documentId,
      path: file,
      mediaType,
      size: fs.statSync(file).size,
      hash,
    });
    return id;
  }
  discard(id, documentId) {
    const c = this.get("candidate", id);
    requireThat(
      c.documentId === documentId,
      "ACCESS_DENIED",
      "Candidate belongs to another document.",
    );
    requireThat(
      !c.published,
      "ALREADY_COMMITTED",
      "Published candidate cannot be discarded.",
    );
    this.put("candidate", id, { ...c, discarded: true });
    if (fs.existsSync(c.path)) fs.unlinkSync(c.path);
    return { discarded: true };
  }
  async validateCandidate(d, c, signal) {
    const stages = [
      ...(c.stages ?? []),
      {
        base: c.validationBase ?? d.snapshot,
        path: c.path,
        revision: c.revision,
        operations: c.operations,
      },
    ];
    let base = d.snapshot;
    const results = [];
    for (const stage of stages) {
      requireThat(
        stage.base === base,
        "INTEGRITY_FAILED",
        "Candidate validation chain is incomplete.",
      );
      const serialized = d.format==='docx' ? boundSerializedValidation(stage.operations,
        this.docxSnapshotRevisions.revision(stage.base,this.config.maxFileBytes),
        this.docxSnapshotRevisions.revision(stage.path,this.config.maxFileBytes)) : null;
      const objects = d.format === "docx" && stage.operations.length > 0 && stage.operations.every(operation => operation.preservationProfile === "docx.guarded-objects.v1");
      const result = serialized ?? await this.worker(
        objects ? "docx_opc_validate" : stage.operations.length === 0
          ? "office_validate_created"
          : d.format + "_validate",
        stage.path,
        { documentId: c.documentId },
        {
          original: stage.base,
          operations: stage.operations,
          format: d.format,
          ...(objects ? { bypassDocx: true } : {}),
        },
        signal,
      );
      requireThat(
        result.passed && result.revision === stage.revision,
        "VALIDATION_FAILED",
        "An editing phase failed preservation validation.",
      );
      results.push(result);
      base = stage.path;
    }
    const result = results.at(-1);
    if (results.length === 1) return result;
    return {
      ...result,
      profile: "office.validated-phases.v1",
      phases: results.map((r) => ({
        profile: r.profile,
        revision: r.revision,
      })),
      checks: [
        ...new Set(results.flatMap((r) => r.checks)),
        "ordered-edit-phase-chain",
      ],
      notVerified: [...new Set(results.flatMap((r) => r.notVerified ?? []))],
    };
  }

  async operation(name, a, signal, options = {}) {
    if (name === "office_commit") return this.documents.publishCandidate(a, signal, options);
    const request = options.request ?? { name, args: a };
    return this.durable.run(
      name,
      a,
      async ({ operation: durable }) => {
        this.checkQuota(this.config.maxFileBytes);
        let result;
        const d = this.doc(a.documentId);
        requireThat(
          name.startsWith(d.format + "_"),
          "UNSUPPORTED_FORMAT",
          "Editing tool must match source format.",
        );
        this.checkSource(d, a.expectedSourceRevision);
        if(d.format==='docx')this.documents.preparePublicationChecks();
        requireThat(
          Boolean(a.candidateId) === Boolean(a.expectedCandidateRevision),
          "INVALID_ARGUMENT",
          "Candidate ID and expected revision must be supplied together.",
        );
        const c = a.candidateId ? this.candidate(a).c : null;
        requireThat(
          !c?.published,
          "ALREADY_COMMITTED",
          "Published candidates are immutable; open the output for another edit.",
        );
        const candidateId = a.candidateId ?? newId("candidate"),
          out = this.internal("candidates", newId("revision") + "." + d.format);
        const transition =
          c &&
          !c.operations.every(
            (op) => op.preservationProfile === a.preservationProfile,
          );
        requireThat(
          !transition || options.transition === true,
          "PROFILE_CONFLICT",
          "Use one preservation profile per candidate; commit and reopen before changing engines.",
        );
        if (transition) await this.validateCandidate(d, c, signal);
        const edited = await this.worker(
          name,
          c?.path ?? d.snapshot,
          a,
          { output: out, ...(options.workerContext ?? {}) },
          signal,
        );
        requireThat(
          !signal?.aborted,
          "CANCELLED",
          "Edit cancelled before accepting the candidate.",
        );
        this.doc(a.documentId);
        const currentCandidate = c ? this.candidate(a).c : null;
        if (currentCandidate)
          requireThat(
            !currentCandidate.published,
            "ALREADY_COMMITTED",
            "Candidate was published during edit.",
          );
        requireThat(
          fs.statSync(out).size <= this.config.maxFileBytes,
          "LIMIT_EXCEEDED",
          "Generated candidate exceeds file size limit.",
        );
        this.checkQuota();
        if(d.format==='docx'&&edited.engineEvidence?.verification==='superdoc-serialized-heading-v1')
          requireThat(this.docxSnapshotRevisions.revision(out,this.config.maxFileBytes)===edited.revision,
            'INTEGRITY_FAILED','Verified serialized candidate changed before acceptance.');
        const operations = [
          ...(transition ? [] : (c?.operations ?? [])),
          {
            ...a,
            tool: name,
            ...(edited.engineEvidence
              ? { engineEvidence: edited.engineEvidence }
              : {}),
          },
        ];
        this.checkpoint("candidate-before-commit");
        this.db.exec("BEGIN IMMEDIATE");
        this.put("candidate", candidateId, {
          documentId: a.documentId,
          path: out,
          revision: edited.revision,
          operations,
          validationBase: transition
            ? c.path
            : (c?.validationBase ?? d.snapshot),
          stages: transition
            ? [
                ...(c.stages ?? []),
                {
                  base: c.validationBase ?? d.snapshot,
                  path: c.path,
                  revision: c.revision,
                  operations: c.operations,
                },
              ]
            : (c?.stages ?? []),
          published: false,
          discarded: false,
        });
        result = {
          operationId: a.operationId,
          data: {
            candidateId,
            candidateRevision: edited.revision,
            sourceRevision: d.revision,
            applied: true,
            validationPassed: false,
            published: false,
            ...(edited.engineEvidence
              ? {
                  engine: edited.engine,
                  evaluation: !!edited.engineEvidence.evaluation,
                  engineEvidence: edited.engineEvidence,
                }
              : {}),
            ...(edited.receipt
              ? { action: edited.action, receipt: edited.receipt }
              : {}),
            affectedTargets: edited.affectedTargets,
            ...(edited.readback ? { readback: edited.readback } : {}),
            ...(edited.workflowVerification ? { workflowVerification: edited.workflowVerification } : {}),
            ...(edited.workflow ? { workflow: edited.workflow } : {}),
            ...(edited.postconditions
              ? {
                  postconditions: edited.postconditions,
                  semanticDelta: edited.semanticDelta,
                }
              : {}),
          },
        };
        if (["xlsx", "csv"].includes(d.format)) {
          result = boundWorkbookMutationEvidence(result, {
            maxResponseBytes: this.config.maxResponseBytes,
            documentId: a.documentId,
          });
        }
        this.durable.complete(durable, result);
        this.db.exec("COMMIT");
        this.checkpoint("candidate-after-commit");
        return result;

      },
      {
        fingerprintName: request.name,
        fingerprintArgs: request.args,
      },
    );
  }
  published(operationId, p) {
    const descriptor=p.provenance?publishArtifactManifest(p.dest,p.provenance):undefined;
    const artifact = p.kind === "artifact";
    const c = artifact ? null : this.get("candidate", p.candidateId);
    if (c) this.put("candidate", p.candidateId, { ...c, published: true });
    const mediaType = artifact
      ? p.mediaType
      : ({
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        csv: "text/csv",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        pdf: "application/pdf",
      })[this.get("document", p.documentId).format];
    const assetId = this.asset(p.dest, p.documentId, mediaType,p.verifiedStageRevision===p.hash?p.hash:undefined);
    return {
      operationId,
      data: {
        published: true,
        outputRevision: p.hash,
        assetId,
        path: p.dest,
        ...(descriptor?{artifactDescriptor:descriptor}:{}),
        ...(artifact
          ? { format: p.format, engine: p.engine, notVerified: p.notVerified ?? [] }
          : {
              candidateId: p.candidateId,
              applied: true,
              validationPassed: true,
              ...(c.operations.some((op) => op.engineEvidence)
                ? { evaluation: c.operations.some((op) => op.engineEvidence?.evaluation) }
                : {}),
            }),
      },
    };
  }
  reconcile(id, op) {
    const p = op.publication;
    try {
      const tmp = fs.statSync(p.tmp),
        out = fs.statSync(p.dest);
      if (
        ((tmp.ino === out.ino && tmp.dev === out.dev) || (p.hosted && p.hostReceipt?.revision === p.hash.replace(/^sha256:/, ""))) &&
        fileRevision(p.dest, this.config.maxFileBytes) === p.hash
      ) {
        syncDirectory(path.dirname(p.dest));
        const result = this.published(id, p);
        this.put("operation", id, { ...op, status: "completed", result });
        try {
          fs.unlinkSync(p.tmp);
          if(p.stageDirectory)fs.rmdirSync(p.stageDirectory);
        } catch {}
        return;
      }
    } catch {}
    this.put("operation", id, {
      ...op,
      status: "failed",
      error: {
        code: "PUBLICATION_UNCERTAIN",
        message:
          "Interrupted publication was not verified. Source is unchanged; inspect destination before new work.",
      },
    });
  }
  recover() {
    for (const op of this.rows("operation"))
      if (op.status === "completed" && op.publication) {
        const p = op.publication;
        try {
          const tmp = fs.lstatSync(p.tmp),
            out = fs.lstatSync(p.dest);
          if (
            tmp.isFile() &&
            out.isFile() &&
            tmp.ino === out.ino &&
            tmp.dev === out.dev &&
            fileRevision(p.dest, this.config.maxFileBytes) === p.hash
          ) {
            fs.unlinkSync(p.tmp);
            if(p.stageDirectory)fs.rmdirSync(p.stageDirectory);
          }
        } catch {} // The durable receipt is never downgraded by temporary-file cleanup.
      }
    for (const op of this.rows("operation"))
      if (op.status === "publishing") this.reconcile(op.id, op);
      else if (op.status === "running")
        this.put("operation", op.id, {
          ...op,
          status: "failed",
          error: {
            code: "PARTIAL_MUTATION",
            message:
              "Service stopped during mutation. Candidate outcome requires inspection; do not repeat with a new operation ID.",
          },
        });
    for (const kind of ["job", "task"])
      for (const record of this.rows(kind))
        if (record.status === "running")
          this.put(kind, record.id, {
            ...record,
            status: "failed",
            ...(kind === "task" ? { cancellation: "closed" } : {}),
            result: {
              error: {
                code: "WORKER_FAILED",
                message: `${kind === "job" ? "Job" : "Task"} interrupted by service restart.`,
              },
            },
          });
  }
}
