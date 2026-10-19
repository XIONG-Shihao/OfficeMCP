import {freshRevisionProof} from '../../infrastructure/streaming-file-revision.mjs';
import { saveChecked, ENGINE } from "./docx-sdk.mjs";
import fs from "node:fs";
import { serializedHeadingStep, captureSerializedBaseline, preflightSerializedBaseline, releaseSerializedBaseline, serializedHeadingEvidence } from "./docx-serialized-heading-proof.mjs";
import {tracePhase} from "../../infrastructure/docx-phase-trace.mjs";
import { capturePostconditionBaseline, verifyPostconditions } from "./docx-agent-postconditions.mjs";

/** One commit/discard boundary for native and agent step adapters. */
export async function executeDocxMutationPlan({ client, doc, task, preview = false, executeStep, discard, setDocument, resetClient, restart, allowSerializedProof=true, userName, verifyWorkflow, captureReadback }) {
  let baseline;
  try {
    const inputRevision = freshRevisionProof(task.input, task.maxFileBytes).revision;
    const headingStep = !preview && allowSerializedProof && serializedHeadingStep(task);
    baseline = headingStep ? await tracePhase('serializedBaseline',()=>captureSerializedBaseline(doc,task)) : null;
    if(baseline&&!await tracePhase('serializedEligibility',()=>preflightSerializedBaseline(task,baseline,inputRevision))) {
      // The original baseline is incompatible with the narrow proof. No edit
      // has run yet: use full semantic verification on this same SDK session.
      fs.rmSync(baseline.file,{force:true});baseline=null;
    }
    const beforeText = await capturePostconditionBaseline(doc, task.args.expect);
    const steps = [];
    for (const [index,step] of task.args.steps.entries()) {
      try {steps.push(await executeStep(step));}
      catch(error){error.message=`Batch step ${index+1} (${step.action??step.type}) failed after ${steps.length} completed step(s); the failed SDK session is discarded. ${error.message}`;throw error;}
    }
    const checks = await verifyPostconditions(doc, task.args.expect, beforeText);
    const workflowVerification = await verifyWorkflow?.();
    const evidence = workflowVerification ? { workflowVerification } : {};
    if (preview) {
      const readback = await captureReadback?.(true);
      await discard();
      return { steps, ...checks, ...evidence, ...(readback ? { readback } : {}) };
    }
    let saved;
    if (baseline) {
      await tracePhase('sdkExport',()=>doc.save({out:task.output,force:false}));
      try {
        const {proof,independent}=await tracePhase('serializedPreservation',async()=>serializedHeadingEvidence(task,baseline,headingStep));
        saved={revision:proof.outputRevision,engine:ENGINE,engineEvidence:{engine:ENGINE,
          verification:'superdoc-serialized-heading-v1',inputRevision,outputRevision:proof.outputRevision,
          serializedProof:proof,packageEvidence:independent,evaluation:false,
          postconditions:['SDK-request-postconditions',...proof.checks,...independent.checks],
          notVerified:['visual-layout']}};
      } catch (error) {
        // Candidate-dependent changes still need the conservative retry. Source
        // baseline normalization has already been checked before the first edit.
        if(!['PRESERVATION_FAILED','UNSUPPORTED_FEATURE','VALIDATION_FAILED'].includes(error.code))throw error;
        if(!restart)throw error;
        releaseSerializedBaseline(baseline);
        const replacement=await restart();
        fs.rmSync(task.output,{force:true});
        return await executeDocxMutationPlan({client:replacement.client,doc:replacement.doc,task,preview,
          executeStep,discard,setDocument,resetClient,restart,allowSerializedProof:false,userName,verifyWorkflow,captureReadback});
      }
    }
    if(!saved) saved = await saveChecked(client, doc, task, inputRevision, { setDocument, resetClient, userName,
      inspectReopened: captureReadback ? reopened => captureReadback(false, reopened) : undefined });
    const readback = saved.readback;
    return { saved, steps, ...checks, ...evidence, ...(readback ? { readback } : {}) };
  } catch (error) {
    // A later failed step must not leave earlier mutations in the reusable host.
    await discard();
    throw error;
  } finally {
    if(baseline){releaseSerializedBaseline(baseline);fs.rmSync(baseline.file,{force:true});}
  }
}
