import { DocumentSessions } from "./document-sessions.mjs";
import fs from "node:fs";
import path from "node:path";
import {Worker} from 'node:worker_threads';
import {pathToFileURL} from 'node:url';
import {fileIdentity} from '../infrastructure/streaming-file-revision.mjs';
import { OfficeError, requireThat, errorPayload } from "../infrastructure/errors.mjs";
import { fileRevision, newId, destinationPath, syncDirectory,sourcePath } from "../infrastructure/files.mjs";
import {tracePhase,phaseEvent} from "../infrastructure/docx-phase-trace.mjs";

/** Candidate evidence and publication policy, independent of public tool routing. */
export class DocumentRuntime extends DocumentSessions {
  constructor(service) {
    super(service);
    this.stagingWorkers=new Set();
    this.publicationSlots=[];this.publicationQueue=[];this.publicationJobs=0;this.publicationSequence=0;this.publicationClosed=false;
  }

  async admitPublication(input) {
    try { return await this.service.publication.admit(input); }
    catch { throw new OfficeError("HOST_PUBLICATION_REFUSED", "Host publication admission failed. Check the active run and manual document edits."); }
  }

  async validate(args, signal) {
    const service = this.service;
    const { d, c } = service.candidate(args);
    const id = newId("validation");
    const record = {
      documentId: args.documentId,
      candidateId: args.candidateId,
      revision: c.revision,
    };
    try {
      const structural = await tracePhase('candidateValidation',()=>service.validateCandidate(d, c, signal));
      service.candidate(args);
      requireThat(!signal?.aborted, "CANCELLED", "Validation cancelled before accepting evidence.");
      requireThat(structural.revision === args.expectedCandidateRevision,
        "CANDIDATE_REVISION_CONFLICT", "Candidate changed during validation.");
      const completed = { ...record, state: "completed", result: structural };
      service.put("validation", id, completed);
      return { id, ...completed };
    } catch (error) {
      service.put("validation", id, { ...record, state: "failed", result: { passed: false }, error: errorPayload(error) });
      throw error;
    }
  }
  async save(a, signal, requestName = "office_save") {
    const service = this.service;
    requireThat(
      a.view.candidateId,
      "INVALID_ARGUMENT",
      "office_save requires a candidate view from office_create/edit.",
    );
    const d = service.doc(a.view.documentId),
      c = service.get("candidate", a.view.candidateId);
    requireThat(
      c.documentId === a.view.documentId &&
        !c.discarded &&
        c.revision === a.view.revision,
      "REVISION_CONFLICT",
      "Candidate view is stale or unavailable.",
    );
    const publicationTicket = d.format === "docx" && service.publication && service.get("operation", a.operationId, false)?.status !== "completed"
      ? await this.admitPublication({ document: d.generated ? null : d.source, destination: destinationPath(service.config, a.destination), expectedRevision: d.revision, candidateRevision: c.revision }) : undefined;
    let preparedStage,preparedSource;
    try {
    let validation = service
      .rows("validation")
      .find(
        (v) =>
          v.candidateId === a.view.candidateId &&
          v.revision === c.revision &&
          v.result.passed,
      );
    // The previous mutation already produced a verified immutable candidate.
    // Validate large files in a fresh SDK worker so old parse/export buffers
    // cannot overlap with the independent semantic/package validation pass.
    const serialized=d.format==='docx'&&c.operations?.length&&c.operations.every(op=>op.engineEvidence?.verification==='superdoc-serialized-heading-v1');
    if (!serialized && d.format === "docx" && fs.statSync(c.path).size > 64 * 1024 * 1024) {
      await service.docxPool.evictIdle(a.view.documentId);
    }
    if (!validation) {
      validation = await this.validate({
        documentId: a.view.documentId,
        candidateId: a.view.candidateId,
        expectedCandidateRevision: a.view.revision,
      }, signal);
    }
    if(serialized){
      service.checkQuota(service.config.maxFileBytes);
      requireThat(this.publicationJobs<=2,'BUSY','Two publication workflows are already active; retry after one completes.');
      preparedSource=this.checkSourceAsync(d,d.revision);
      preparedStage={promise:this.stagePublicationAsync(c.path,destinationPath(service.config,a.destination),c.revision),owned:false};
    }
    // The candidate is already serialized and validated. Large SDK sessions retain
    // private copies that publication no longer needs; release them before quota checks.
    if (d.format === "docx" && fs.statSync(c.path).size > 64 * 1024 * 1024) {
      await service.docxPool.evictIdle(a.view.documentId);
      service.quotaSnapshot.checkedAt = Number.NEGATIVE_INFINITY;
    }
    const result = await this.publishCandidate(
      {
        documentId: a.view.documentId,
        candidateId: a.view.candidateId,
        expectedCandidateRevision: a.view.revision,
        expectedSourceRevision: d.revision,
        operationId: a.operationId,
        validationId: validation.id,
        mode: "save_as",
        overwrite: false,
        destination: a.destination,
      },
      signal,
      { request: { name: requestName, args: a }, publicationTicket,preparedStage,preparedSource },
    );
    return {
      ...result,
      view: a.view,
      validation: {
        checks: validation.result.checks,
        notVerified: validation.result.notVerified,
        profile: validation.result.profile,
      },
    };
    } finally {
      if(preparedSource)await preparedSource.catch(()=>{});
      if(preparedStage&&!preparedStage.owned){const stage=await preparedStage.promise.catch(()=>null);if(stage){fs.rmSync(stage.tmp,{force:true});if(stage.stageDirectory)fs.rmSync(stage.stageDirectory,{recursive:true,force:true});}}
      if (publicationTicket) await service.publication.cancel(publicationTicket).catch(() => {});
    }
  }

  async publishCandidate(a, signal, options = {}) {
    const service = this.service;
    const request = options.request ?? { name: "office_commit", args: a };
    const source = service.doc(a.documentId), candidate = service.get("candidate", a.candidateId);
    const publicationTicket = options.publicationTicket ?? (source.format === "docx" && service.publication && service.get("operation", a.operationId, false)?.status !== "completed"
      ? await this.admitPublication({ document: source.generated ? null : source.source, destination: destinationPath(service.config, a.destination), expectedRevision: source.revision, candidateRevision: candidate.revision }) : undefined);
    try { return await service.durable.run("office_commit", a, async ({ operation: durable }) => {
      service.checkQuota(service.config.maxFileBytes);
      requireThat(!signal?.aborted, "CANCELLED", "Publication cancelled before saving.");
      const { d, c } = service.candidate(a);
      requireThat(
        !c.published,
        "ALREADY_COMMITTED",
        "Candidate has already been published.",
      );
      const v = service.get("validation", a.validationId);
      requireThat(
        v.candidateId === a.candidateId &&
          v.documentId === a.documentId &&
          v.revision === c.revision &&
          v.result.passed,
        "VALIDATION_FAILED",
        "Validation does not bind this exact candidate.",
      );
      const dest = destinationPath(service.config, a.destination);
      requireThat(
        !fs.existsSync(dest),
        "DESTINATION_CONFLICT",
        "Destination already exists; save-as never overwrites.",
      );
      requireThat(
        path.extname(dest).toLowerCase() === "." + d.format,
        "INVALID_ARGUMENT",
        "Output extension must match the source format.",
      );
      // Stage an independent private copy while the coordinator hashes the live
      // source. Both checks finish before any publication or durable intent.
      let preparedStage=options.preparedStage;
      const serialized=d.format==='docx'&&c.operations?.length&&c.operations.every(op=>op.engineEvidence?.verification==='superdoc-serialized-heading-v1');
      if(serialized&&!preparedStage)preparedStage={promise:this.stagePublicationAsync(c.path,dest,c.revision),owned:false};
      let sourceGuard;
      try { sourceGuard=options.preparedSource?await options.preparedSource:service.checkSource(d, a.expectedSourceRevision); }
      catch(error){if(preparedStage){const stage=await preparedStage.promise.catch(()=>null);if(stage){fs.rmSync(stage.tmp,{force:true});if(stage.stageDirectory)fs.rmSync(stage.stageDirectory,{recursive:true,force:true});}}throw error;}
      return this.finishPublication(durable, a.operationId, {
        dest, file: c.path, hash: c.revision, candidateId: a.candidateId, documentId: a.documentId,
        ...(d.format === "docx" && service.publication ? { hosted: true, source: d.generated ? null : d.source, sourceRevision: d.revision, publicationTicket } : {}),
      },preparedStage,signal,sourceGuard);
    }, { fingerprintName: request.name, fingerprintArgs: request.args });
    } finally { if (publicationTicket && !options.publicationTicket) await service.publication.cancel(publicationTicket).catch(() => {}); }
  }

  stagePublication(file, destination, expectedRevision) {
    const tmp = path.join(path.dirname(destination), ".officemcp-" + newId("publish") + ".tmp");
    try {
      // Copy beside the destination so the final atomic link always stays on one filesystem.
      // FICLONE is opportunistic; copyFile falls back to a regular copy where unavailable.
      fs.copyFileSync(file, tmp, fs.constants.COPYFILE_EXCL | fs.constants.COPYFILE_FICLONE);
      fs.chmodSync(tmp, 0o600);
      requireThat(fileRevision(tmp, this.service.config.maxFileBytes) === expectedRevision,
        "INTEGRITY_FAILED", "Publication artifact changed before staging.");
      const fd = fs.openSync(tmp, "r");
      try { fs.fsyncSync(fd); } finally { fs.closeSync(fd); }
      syncDirectory(path.dirname(tmp));
      return tmp;
    } catch (error) {
      try { fs.unlinkSync(tmp); } catch {}
      throw error;
    }
  }

  stagePublicationAsync(file,destination,expectedRevision) {
    const stageDirectory=fs.mkdtempSync(path.join(path.dirname(destination),'.officemcp-stage-'));
    fs.chmodSync(stageDirectory,0o700);
    const tmp=path.join(stageDirectory,'candidate.docx');
    try{return this.publicationWorker({file,tmp,stageDirectory,expectedRevision,maxBytes:this.service.config.maxFileBytes});}
    catch(error){fs.rmSync(stageDirectory,{recursive:true,force:true});throw error;}
  }
  checkSourceAsync(d,expectedRevision){
    requireThat(expectedRevision===d.revision,'SOURCE_REVISION_CONFLICT','Source revision is stale.');
    const file=d.generated?d.snapshot:sourcePath(this.service.config,d.source);
    return this.publicationWorker({operation:'source',file,expectedRevision,maxBytes:this.service.config.maxFileBytes});
  }
  preparePublicationChecks(){
    requireThat(!this.publicationClosed&&!this.service.closed,'CANCELLED','Publication checker is closed.');
    while(this.publicationSlots.length<2){
      requireThat(this.stagingWorkers.size<4,'BUSY','Publication checkers are still draining.');
      const worker=new Worker(new URL('../infrastructure/publication-stage-worker.mjs',import.meta.url),{execArgv:[]});
      const slot={worker,pending:null,dead:false};this.publicationSlots.push(slot);this.stagingWorkers.add(worker);worker.unref();
      const finish=(error,result)=>{
        const job=slot.pending;if(!job)return;
        slot.pending=null;clearTimeout(job.timer);this.publicationJobs--;worker.unref();
        const failure=job.abandoned??error;
        if(failure){if(job.data.tmp)fs.rmSync(job.data.tmp,{force:true});if(job.data.stageDirectory)fs.rmSync(job.data.stageDirectory,{recursive:true,force:true});job.reject(failure);}else job.resolve(result);
        this.pumpPublicationChecks();
      };
      worker.on('message',message=>{
        if(message.id!==slot.pending?.data.id){finish(new OfficeError('WORKER_FAILED','Publication response identity differs.'));return;}
        finish(message.error?new OfficeError(slot.pending.data.operation==='source'?'SOURCE_REVISION_CONFLICT':'INTEGRITY_FAILED',message.error):null,message.result);
      });
      worker.once('error',error=>{slot.dead=true;this.publicationSlots=this.publicationSlots.filter(s=>s!==slot);finish(error)});
      worker.once('exit',code=>{
        this.stagingWorkers.delete(worker);this.publicationSlots=this.publicationSlots.filter(s=>s!==slot);
        finish(new OfficeError('WORKER_FAILED',`Publication checker exited ${code}.`));
      });
      worker.unref();
    }
  }
  pumpPublicationChecks(){
    if(this.publicationClosed||this.service.closed)return;
    for(const slot of this.publicationSlots){
      if(slot.pending||!this.publicationQueue.length)continue;
      slot.pending=this.publicationQueue.shift();slot.worker.ref();
      const job=slot.pending;
      job.timer=setTimeout(()=>{slot.dead=true;job.abandoned=new OfficeError('WORKER_TIMEOUT','Publication checker exceeded its deadline.');job.reject(job.abandoned);this.publicationSlots=this.publicationSlots.filter(s=>s!==slot);void slot.worker.terminate().catch(()=>{});},this.service.config.workerTimeoutMs??60000);
      slot.worker.postMessage(job.data);
    }
  }
  publicationWorker(data){
    requireThat(this.publicationJobs<4,'BUSY','Four bounded publication checks are already active.');
    this.preparePublicationChecks();
    const promise=new Promise((resolve,reject)=>{
      this.publicationJobs++;this.publicationQueue.push({data:{...data,id:++this.publicationSequence},resolve,reject});this.pumpPublicationChecks();
    });
    void promise.catch(()=>{});return promise;
  }
  closePublicationChecks(){
    this.publicationClosed=true;
    for(const job of this.publicationQueue.splice(0)){this.publicationJobs--;job.reject(new OfficeError('CANCELLED','Publication checker closed.'));}
    for(const worker of this.stagingWorkers)void worker.terminate().catch(()=>{});
  }

  async finishPublication(durable, operationId, artifact,preparedStage,signal,sourceGuard) {
    const publicationStart=performance.now();
    const service = this.service;
    const document=service.doc(artifact.documentId);
    artifact.provenance={schema:'mcp-artifact/1',sha256:artifact.hash.replace(/^sha256:/,''),
      inputs:document.generated?[]:[{uri:pathToFileURL(sourcePath(service.config,document.source)).href,sha256:document.revision.replace(/^sha256:/,'')}]};
    const { file, publicationTicket, ...metadata } = artifact;
    const ticket = artifact.hosted ? publicationTicket ?? await this.admitPublication({ document: artifact.source, destination: artifact.dest, expectedRevision: artifact.sourceRevision, candidateRevision: artifact.hash }) : undefined;
    let publication;
    try {
    const stage=await tracePhase('staging',()=>preparedStage?preparedStage.promise:this.stagePublication(file, artifact.dest, artifact.hash));
    publication = { ...metadata,tmp:typeof stage==='string'?stage:stage.tmp,
      ...(typeof stage==='string'?{}:{stageIdentity:stage.identity,stageDirectory:stage.stageDirectory,verifiedStageRevision:artifact.hash}) };
    try { requireThat(!signal?.aborted&&!service.closed,'CANCELLED','Publication cancelled before committing its staged file.');service.durable.publishing(durable, publication);if(preparedStage)preparedStage.owned=true; }
    catch (error) {
      // Before journaling/linking this file has no recovery owner and is safe to remove.
      try { fs.unlinkSync(publication.tmp); } catch {}
      if(publication.stageDirectory)fs.rmSync(publication.stageDirectory,{recursive:true,force:true});
      throw error;
    }
    service.checkpoint("publication-before-link");
    try {
      if(sourceGuard){
        const stat=fs.lstatSync(sourceGuard.file,{bigint:true});
        if(!stat.isFile()||JSON.stringify(fileIdentity(stat))!==JSON.stringify(sourceGuard.identity)){
          const error=new OfficeError('SOURCE_REVISION_CONFLICT','Source changed during publication staging.');error.notPublished=true;throw error;
        }
      }
      if(publication.stageIdentity){
        const stat=fs.lstatSync(publication.tmp,{bigint:true});
        if(!stat.isFile()||JSON.stringify(fileIdentity(stat))!==JSON.stringify(publication.stageIdentity)){
          const error=new OfficeError('INTEGRITY_FAILED','Verified publication stage changed before linking.');error.notPublished=true;throw error;
        }
      }
      if (ticket) {
        const receipt = await service.publication.publish(ticket, publication.tmp);
        requireThat(receipt.revision === artifact.hash.replace(/^sha256:/, ""), "INTEGRITY_FAILED", "Host publication receipt differs from the validated candidate.");
        publication.hostReceipt = receipt;
        service.durable.publishing(durable, publication);
      } else {
        fs.linkSync(publication.tmp, publication.dest);
        if(publication.stageIdentity){
          const stat=fs.lstatSync(publication.dest,{bigint:true}),id=fileIdentity(stat);
          requireThat(stat.isFile()&&['dev','ino','size','mtimeNs'].every(key=>id[key]===publication.stageIdentity[key]),
            'INTEGRITY_FAILED','Publication destination differs from the verified stage.');
        }
      }
    } catch (error) {
      if (error.notPublished === true) {
        try { fs.unlinkSync(publication.tmp); } catch {}
        if(publication.stageDirectory)fs.rmSync(publication.stageDirectory,{recursive:true,force:true});
        const refusal = new OfficeError("HOST_PUBLICATION_REFUSED", error.message);
        service.durable.fail(durable, refusal);
        throw refusal;
      }
      if (error.code === "EEXIST") {
        const conflict = new OfficeError("DESTINATION_CONFLICT", "Destination was created concurrently.");
        // EEXIST proves our link was not installed; do not turn a known collision into uncertainty.
        try { fs.unlinkSync(publication.tmp); } catch {}
        if(publication.stageDirectory)fs.rmSync(publication.stageDirectory,{recursive:true,force:true});
        service.durable.fail(durable, conflict);
        throw conflict;
      }
      throw error;
    }
    service.checkpoint("publication-after-link");
    syncDirectory(path.dirname(publication.dest));
    const result = service.published(operationId, publication);
    service.durable.complete(durable, result, publication);
    service.checkpoint("publication-after-receipt");
    try { fs.unlinkSync(publication.tmp); } catch {}
    if(publication.stageDirectory)fs.rmSync(publication.stageDirectory,{recursive:true,force:true});
    return result;
    } finally { if (ticket) await service.publication.cancel(ticket).catch(() => {});phaseEvent('publication',publicationStart); }
  }

  async exportArtifact(name, args, format, produce, signal) {
    const service = this.service;
    return service.durable.run(name, args, async ({ operation: durable }) => {
      const checkView = () => {
        service.view({ documentId: args.view.documentId,
          candidateId: args.view.candidateId, expectedRevision: args.view.revision });
      };
      checkView();
      const dest = destinationPath(service.config, args.destination);
      requireThat(path.extname(dest).toLowerCase() === "." + format,
        "INVALID_ARGUMENT", `Export destination must use the .${format} extension.`);
      requireThat(!fs.existsSync(dest), "DESTINATION_CONFLICT", "Destination already exists; export never overwrites.");
      const artifact = await produce();
      checkView();
      requireThat(!signal?.aborted, "CANCELLED", "Export cancelled before publication.");
      const result = await this.finishPublication(durable, args.operationId, {
        kind: "artifact", format, mediaType: artifact.mediaType,
        dest, file: artifact.file, hash: fileRevision(artifact.file, service.config.maxFileBytes), documentId: args.view.documentId,
        engine: artifact.engine, notVerified: artifact.notVerified,
        ...(format === "docx" && service.publication ? {
          hosted: true,
          source: service.doc(args.view.documentId).format === "docx" && !service.doc(args.view.documentId).generated ? service.doc(args.view.documentId).source : null,
          sourceRevision: service.doc(args.view.documentId).revision,
        } : {}),
      });
      try { fs.unlinkSync(artifact.file); } catch {}
      return result;
    });
  }
}
