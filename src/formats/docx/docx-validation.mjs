import fs from "node:fs";
import { packageEvidence } from "../opc/engine-evidence.mjs";
import { digest, fingerprint } from "../../infrastructure/files.mjs";
import { requireThat } from "../../infrastructure/errors.mjs";
import { snapshot } from "./docx-sdk.mjs";
import {boundSerializedValidation} from './docx-serialized-heading-proof.mjs';

export async function validateDocxCandidate(doc, task) {
  const bytes = fs.readFileSync(task.input);
  const originalBytes=fs.readFileSync(task.original);
  const serialized=boundSerializedValidation(task.operations,digest(originalBytes),digest(bytes));
  if(serialized)return serialized;
  let revision = digest(originalBytes);
  for (const operation of task.operations ?? []) {
    requireThat(
      operation.engineEvidence?.verification === "superdoc-sdk-readback-v1" &&
        operation.engineEvidence.inputRevision === revision,
      "VALIDATION_FAILED",
      "This candidate predates the current SuperDoc and package-validation contract. Reopen the saved source and edit again.",
    );
    revision = operation.engineEvidence.outputRevision;
  }
  requireThat(
    revision === digest(bytes),
    "VALIDATION_FAILED",
    "Candidate hash chain differs.",
  );
  const last = task.operations?.at(-1);
  if(last) {
    // The operation ledger is host-owned. saveChecked already reopened and
    // compared the exact file whose digest terminates the checked chain.
    // Identical immutable bytes retain that SDK proof; final validation still
    // independently compares the package and the source/candidate revisions.
    requireThat(/^sha256:[a-f0-9]{64}$/.test(last.engineEvidence.sdkSnapshotHash??''),
      "VALIDATION_FAILED","Candidate has no bound SDK semantic readback proof.");
  } else {
    // Creation/legacy inputs without a prior roundtrip proof still require the
    // public SDK to inspect their actual candidate model.
    fingerprint(await snapshot(doc,false));
  }
  const independent = packageEvidence(
    originalBytes,
    bytes,
    task.maxExpandedBytes,
    {
      nativeInflate:true,
      mediaChanges: (task.operations ?? []).flatMap((operation) =>
        (operation.steps ?? [{ action: operation.action }])
          .filter((step) =>
            ["insert_image", "replace_image", "delete_image"].includes(
              step.action,
            ),
          )
          .map((step) => ({
            action: step.action,
            partName: step.targetLocator?.partName,
            replacementHash: step.image?.hash,
          })),
      ),
    },
  );
  return {
    passed: true,
    revision: digest(bytes),
    profile: "docx.sdk-only.v1",
    checks: [
      "SDK-reopen",
      "candidate-hash-chain",
      "SDK-semantic-snapshot",
      ...(last?["SDK-readback-proof-bound-to-exact-candidate-hash"]:[]),
      ...independent.checks,
    ],
    evidence: independent,
    notVerified: [
      ...(!last?.engineEvidence.includeBookmarks
        ? ["untouched-bookmark-metadata"]
        : []),
      "content-not-exposed-by-SDK",
      "visual-layout",
    ],
  };
}
