import { fork } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { OfficeError, requireThat } from "../../infrastructure/errors.mjs";
import { writeDiagnostic } from "../../infrastructure/diagnostics.mjs";
import {workerEnvironment,workerExecArgv} from '../../infrastructure/worker-environment.mjs';
import { ProcessBudget, readProcessTree } from "../../infrastructure/docx-process-budget.mjs";
import {phaseEvent} from "../../infrastructure/docx-phase-trace.mjs";

/**
 * One sequential actor per document, with a bounded global actor semaphore.
 * Independent documents may run concurrently; one document/candidate never does.
 */
export class DocxPool {
  constructor(service) {
    this.service = service;
    this.actors = new Map();
    this.closed = false;
    this.activeRuns = 0;
    this.waiters = [];
    this.maxActors = service.config.maxDocxActors;
    this.maxResident = service.config.maxDocxResidentActors ?? this.maxActors;
    this.pendingRequests = 0;
    this.workers = new Set();
    this.admission = Promise.resolve();
    this.memoryError = null;
    this.changed = new Set();
    this.shutdown = new AbortController();
    this.budget = new ProcessBudget({roots:()=>[...this.workers].map(slot=>slot.child.pid).filter(Boolean),
      groups:()=>process.platform==='win32'?[]:[...this.workers].map(slot=>slot.child.pid).filter(Boolean),
      maximumBytes:service.config.maxDocxProcessBytes ?? 8*1024**3,onExceeded:error=>{
        this.memoryError=new OfficeError("LIMIT_EXCEEDED",`DOCX process memory could not remain within its budget: ${error.message}`);
        for(const actor of this.actors.values()){this.resetActor(actor,this.memoryError);actor.abort.abort(this.memoryError);}
        for(const slot of this.workers)this.service.killWorker(slot.child);
        this.notify();
      }});
    this.budget.start();
  }

  stats() {return {active:this.activeRuns,pending:this.pendingRequests,resident:this.workers.size,waiting:this.waiters.length,
    rssBytes:this.budget.last.rssBytes,maximumObservedBytes:this.budget.maximumObservedBytes};}
  notify(){for(const resolve of this.changed)resolve();this.changed.clear();}
  async waitForRoom(signal) {
    requireThat(!this.closed&&!signal?.aborted,"CANCELLED","DOCX request cancelled.");
    await new Promise((resolve,reject)=>{
      const complete=()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort);this.changed.delete(complete);resolve();};
      const abort=()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort);this.changed.delete(complete);reject(new OfficeError("CANCELLED","Queued DOCX request cancelled."));};
      const timer=setTimeout(()=>{clearTimeout(timer);signal?.removeEventListener('abort',abort);this.changed.delete(complete);reject(new OfficeError("WORKER_TIMEOUT","DOCX resident admission deadline exceeded."));},this.service.config.workerTimeoutMs);
      signal?.addEventListener('abort',abort,{once:true});this.changed.add(complete);
    });
  }
  async releaseWorker(slot) {
    if(slot.reaping)return slot.reaping;
    slot.reaping=(async()=>{
      // Detached groups cover SDK descendants briefly orphaned after the Node
      // worker exits. Do not release a resident reservation until all are gone.
      if(slot.child.pid)while(true){const tree=await readProcessTree([slot.child.pid],{groups:process.platform==='win32'?[]:[slot.child.pid]});if(!tree.processes.length)break;await new Promise(r=>setTimeout(r,20));}
      this.workers.delete(slot);if(!this.workers.size)this.memoryError=null;this.notify();
    })().catch(error=>{this.memoryError=new OfficeError("WORKER_FAILED",`Unable to confirm DOCX worker cleanup: ${error.message}`);this.notify();});
    return slot.reaping;
  }

  get slot() {
    const slots = [...this.actors.values()].map((actor) => actor.slot).filter(Boolean);
    return slots.find((slot) => slot.busy) ?? slots[0];
  }

  key(input, args) {
    return args.documentId ?? args.view?.documentId ?? input ?? "create";
  }

  actor(key) {
    let actor = this.actors.get(key);
    if(actor?.abort.signal.aborted&&!actor.queued){this.actors.delete(key);actor=null;}
    if (!actor) {
      actor = { key, slot: null, tail: Promise.resolve(), queued: 0, abort:new AbortController() };
      this.actors.set(key, actor);
    }
    return actor;
  }

  resetActor(actor, error) {
    const slot = actor?.slot;
    if (!slot) return;
    actor.slot = null;
    clearTimeout(slot.idle);
    if(error)slot.fail?.(error);else slot.cancel?.();
    this.service.killWorker(slot.child);
    this.notify();
  }

  reset(key) {
    if (key !== undefined) {
      const actor = this.actors.get(key);
      actor?.abort.abort(new OfficeError("CANCELLED","DOCX document session reset."));
      this.resetActor(actor);
      this.actors.delete(key);
      return;
    }
    for (const actor of this.actors.values()) {actor.abort.abort(new OfficeError("CANCELLED","DOCX sessions reset."));this.resetActor(actor);}
    this.actors.clear();
  }

  async evictIdle(key, {allowQueued=false}={}) {
    const actor = this.actors.get(key);
    const slot = actor?.slot;
    if (!slot || slot.busy || (actor.queued&&!allowQueued)) return false;
    const exited = slot.child.exitCode !== null || slot.child.signalCode !== null
      ? Promise.resolve()
      : new Promise(resolve => slot.child.once("exit", resolve));
    this.resetActor(actor);
    if(!actor.queued)this.actors.delete(key);
    await exited;
    await this.releaseWorker(slot);
    return true;
  }

  async withResidentAdmission(operation,signal) {
    const previous=this.admission;let release,acquired=false;
    this.admission=new Promise(resolve=>{release=resolve;});
    try {await this.waitForActor(previous,signal);acquired=true;return await operation();}
    finally{if(acquired)release();else previous.then(release);}
  }

  close() {
    this.closed = true;
    this.shutdown.abort(new OfficeError("CANCELLED","DOCX pool closed."));
    this.budget.stop();
    this.reset();
    for (const waiter of this.waiters.splice(0)) {
      clearTimeout(waiter.timer);waiter.signal?.removeEventListener("abort",waiter.abort);
      waiter.reject(new OfficeError("CANCELLED", "DOCX pool closed."));
    }
    this.notify();
  }

  acquire(signal) {
    if(this.closed||signal?.aborted)return Promise.reject(new OfficeError("CANCELLED","Queued DOCX request cancelled."));
    if (this.activeRuns < this.maxActors) {
      this.activeRuns++;
      return Promise.resolve();
    }
    return new Promise((resolve, reject) => {
      const waiter = { resolve, reject, signal, abort: null };
      waiter.abort = () => {
        clearTimeout(waiter.timer);signal?.removeEventListener("abort",waiter.abort);
        const index = this.waiters.indexOf(waiter);
        if (index >= 0) this.waiters.splice(index, 1);
        reject(signal?.reason instanceof OfficeError?signal.reason:new OfficeError("CANCELLED", "Queued DOCX request cancelled."));
      };
      signal?.addEventListener("abort", waiter.abort, { once: true });
      waiter.timer=setTimeout(()=>{waiter.abort();},this.service.config.workerTimeoutMs);
      this.waiters.push(waiter);
    });
  }

  release() {
    this.activeRuns--;
    while (this.waiters.length) {
      const waiter = this.waiters.shift();
      waiter.signal?.removeEventListener("abort", waiter.abort);
      clearTimeout(waiter.timer);
      if (waiter.signal?.aborted) continue;
      this.activeRuns++;
      waiter.resolve();
      return;
    }
  }

  async run(kind, input, args, other, signal) {
    const queuedAt=performance.now();
    requireThat(!this.closed&&!signal?.aborted,"CANCELLED","DOCX pool closed or request cancelled.");
    requireThat(this.pendingRequests<(this.service.config.maxDocxQueuedRequests??32),"BUSY","Global DOCX request queue is full.");
    const actor = this.actor(this.key(input, args));
    signal=AbortSignal.any([this.shutdown.signal,actor.abort.signal,...(signal?[signal]:[])]);
    requireThat(
      actor.queued < 8,
      "BUSY",
      "This document's DOCX request queue is full.",
    );
    actor.queued++;
    this.pendingRequests++;
    const previous = actor.tail;
    let releaseActor;
    actor.tail = new Promise((resolve) => (releaseActor = resolve));
    let acquiredActor = false;
    let acquiredGlobal = false;
    try {
      await this.waitForActor(previous, signal);
      acquiredActor = true;
      await this.acquire(signal);
      acquiredGlobal = true;
      phaseEvent('queueDelay',queuedAt);
      return await this.runNow(actor, kind, input, args, other, signal);
    } finally {
      actor.queued--;
      this.pendingRequests--;
      if (acquiredGlobal) this.release();
      if (acquiredActor) releaseActor();
      else previous.then(releaseActor);
      if(!actor.queued&&!actor.slot&&this.actors.get(actor.key)===actor)this.actors.delete(actor.key);
      this.notify();
    }
  }

  waitForActor(previous, signal) {
    return new Promise((resolve, reject) => {
      const abort = () => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", abort);
        reject(signal?.reason instanceof OfficeError?signal.reason:new OfficeError("CANCELLED", "Queued DOCX request cancelled."));
      };
      const timer = setTimeout(() => {
        signal?.removeEventListener("abort", abort);
        reject(new OfficeError("WORKER_TIMEOUT", "DOCX queue deadline exceeded."));
      }, this.service.config.workerTimeoutMs);
      if (signal?.aborted) return abort();
      signal?.addEventListener("abort", abort, { once: true });
      previous.then(() => {
        clearTimeout(timer);
        signal?.removeEventListener("abort", abort);
        resolve();
      });
    });
  }

  async runNow(actor, kind, input, args, other, signal) {
    const service = this.service;
    requireThat(
      !this.closed && !signal?.aborted,
      "CANCELLED",
      "DOCX actor is closed or request cancelled.",
    );
    requireThat(
      !actor.slot?.busy,
      "BUSY",
      "This DOCX document is processing another request.",
    );
    if (!actor.slot) {
      await this.withResidentAdmission(async()=>{
      while(this.workers.size>=this.maxResident) {
        requireThat(!this.closed&&!signal?.aborted,"CANCELLED","DOCX resident admission cancelled.");
        if(this.memoryError)throw this.memoryError;
        const idle=[...this.actors.values()].filter(a=>a!==actor&&a.slot&&!a.slot.busy).sort((a,b)=>(a.slot.lastUsed??0)-(b.slot.lastUsed??0))[0];
        if(idle)await this.evictIdle(idle.key,{allowQueued:true});else await this.waitForRoom(signal);
      }
      if(this.memoryError)throw this.memoryError;
      await this.budget.check();
      requireThat(!this.closed&&!signal?.aborted,"CANCELLED","DOCX resident admission cancelled.");
      const scratch = fs.mkdtempSync(
          path.join(service.config.stateDirectory, "docx-host-"),
        ),
        child = fork(new URL("./docx-host.mjs", import.meta.url), [], {
          stdio: ["ignore", "ignore", "pipe", "ipc"],
          execArgv: workerExecArgv({heapLimitMb: 768}),
          detached: process.platform !== "win32",
          env: {
            ...workerEnvironment(),
            ...(process.env.RPF_DOCX_MCP_PHASE_TRACE==='1'?{RPF_DOCX_MCP_PHASE_TRACE:'1'}:{}),
          },
        });
      const slot = { child, scratch, busy: false, stderr: "", phaseBuffer:"", lastUsed:Date.now() };
      actor.slot = slot;
      this.workers.add(slot);
      child.stderr.setEncoding("utf8");
      child.stderr.on("data", (chunk) => {
        if(process.env.RPF_DOCX_MCP_PHASE_TRACE==='1'){
          slot.phaseBuffer+=String(chunk);const lines=slot.phaseBuffer.split('\n');slot.phaseBuffer=lines.pop();
          for(const line of lines)if(line.startsWith('[docx-phase] '))process.stderr.write(line+'\n');
        }
        if (slot.stderr.length < 8192)
          slot.stderr += String(chunk).slice(0, 8192 - slot.stderr.length);
      });
      child.on("error", () => {
        slot.fail?.(new OfficeError("WORKER_FAILED", "DOCX host failed."));
        if(!child.pid){void this.releaseWorker(slot);fs.rmSync(scratch,{recursive:true,force:true});}
      });
      child.on("exit", () => {
        if (actor.slot === slot) actor.slot = null;
        clearTimeout(slot.idle);
        const diagnosticId = slot.stderr
          ? writeDiagnostic(
              service.config,
              "docx-host-stderr",
              new Error(slot.stderr),
            )
          : undefined;
        slot.fail?.(
          new OfficeError(
            "WORKER_FAILED",
            diagnosticId
              ? `DOCX host exited. Diagnostic ID: ${diagnosticId}.`
              : "DOCX host exited before returning structured diagnostics.",
          ),
        );
        service.killWorker(child);
        fs.rmSync(scratch, { recursive: true, force: true });
        void this.releaseWorker(slot);
      });
      },signal);
    }
    const slot = actor.slot;
    clearTimeout(slot.idle);
    slot.busy = true;
    service.activeWorkers.add(slot.child);
    return await new Promise((resolve, reject) => {
      let done = false;
      const finish = (error, result) => {
        if (done) return;
        done = true;
        clearTimeout(timer);
        signal?.removeEventListener("abort", cancel);
        slot.child.removeListener("message", message);
        slot.cancel = null;
        slot.fail = null;
        slot.busy = false;
        slot.lastUsed = Date.now();
        service.activeWorkers.delete(slot.child);
        if (error) {
          if (actor.slot === slot) actor.slot = null;
          service.killWorker(slot.child);
          reject(error);
        } else {
          slot.idle = setTimeout(() => {
            if (actor.slot === slot && !slot.busy && !actor.queued) {
              this.resetActor(actor);
              this.actors.delete(actor.key);
            }
          }, service.config.docxIdleMs??60000);
          slot.idle.unref();
          resolve(result);
        }
        this.notify();
      };
      const cancel = () =>
        finish(
          new OfficeError(
            "CANCELLED",
            "DOCX request cancelled; no candidate accepted.",
          ),
        );
      slot.cancel = cancel;
      slot.fail = (error) => finish(error);
      const message = async (response) => {
        try {
          await this.budget.check();
          if(this.memoryError)throw this.memoryError;
          response.ok?finish(null,response.result):finish(new OfficeError(response.error.code,response.error.message));
        }catch(error){finish(error instanceof OfficeError?error:new OfficeError("LIMIT_EXCEEDED",`DOCX memory observation failed: ${error.message}`));}
      };
      const timer = setTimeout(
        () =>
          finish(
            new OfficeError(
              "WORKER_TIMEOUT",
              "DOCX request exceeded its deadline.",
            ),
          ),
        service.config.workerTimeoutMs,
      );
      signal?.addEventListener("abort", cancel, { once: true });
      slot.child.once("message", message);
      slot.child.send(
        {
          kind,
          input,
          args,
          ...other,
          maxExpandedBytes: service.config.maxExpandedBytes,
          maxFileBytes: service.config.maxFileBytes,
          timeoutMs: service.config.workerTimeoutMs,
          scratch: slot.scratch,
          diagnosticConfig: {
            stateDirectory: service.config.stateDirectory,
            readRoots: service.config.readRoots,
            writeRoots: service.config.writeRoots,
          },
        },
        (error) => {
          if (error)
            finish(
              new OfficeError("WORKER_FAILED", "DOCX IPC request failed."),
            );
        },
      );
    });
  }
}
