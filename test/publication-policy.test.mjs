import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createServiceFixture } from './service-fixture.mjs';
import { OfficeService } from '../src/application/service.mjs';

async function candidate(s) {
  return (await s.call('docx_create', { operationId: 'create', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'The service price is 125.' }] } })).view;
}
const binding = view => ({ documentId: view.documentId, candidateId: view.candidateId, expectedCandidateRevision: view.revision });
const save = (view, operationId) => ({ view, operationId, destination: { rootId: 'output', relativePath: `${operationId}.docx` } });

function seedRetiredState(service, view, state, passed = false) {
  service.put('meta', 'schema', { version: 2 });
  service.delete('meta', 'migration-3');
  service.put('candidate', view.candidateId, { ...service.get('candidate', view.candidateId),
    validationRequirement: { revision: view.revision, validationId: 'visual-receipt' } });
  service.put('validation', 'visual-receipt', { ...binding(view), revision: view.revision,
    state, result: { passed }, visualRequest: { pages: [1], expect: { samePageCount: true } },
    visualEvidence: { acceptance: { passed } } });
  const layoutPdf = service.internal('assets', 'legacy-layout.pdf');
  fs.writeFileSync(layoutPdf, 'legacy private layout');
  service.put('docxLayout', 'layout-old', { documentId: view.documentId, path: layoutPdf });
  service.put('asset', 'layout-asset', { documentId: view.documentId, path: layoutPdf });
  return layoutPdf;
}

test('state migration retires pending, failed and passing visual requirements; save runs semantic validation', async () => {
  for (const state of ['pending', 'failed', 'completed']) {
    const s = createServiceFixture({ toolset: 'full' });
    let service = s.service;
    try {
      const view = await candidate(s);
      const layoutPdf = seedRetiredState(service, view, state, state === 'completed');
      const config = service.config;
      service.close();
      service = new OfficeService(config);
      assert.equal(service.stateInfo().schemaVersion, 3);
      assert.equal(service.get('candidate', view.candidateId).validationRequirement, undefined);
      assert.equal(service.get('validation', 'visual-receipt').state, 'retired');
      assert.equal(service.get('validation', 'visual-receipt').result.passed, false);
      assert.equal(service.get('validation', 'visual-receipt').visualEvidence, undefined);
      assert.deepEqual(service.rows('docxLayout'), []);
      assert.equal(service.get('asset', 'layout-asset', false), null);
      assert.equal(service.get('meta', 'migration-3').retiredRequirements, 1);
      const oldReceipt = await service.call('office_commit', { ...binding(view),
        expectedSourceRevision: service.doc(view.documentId).revision,
        operationId: 'old-receipt', validationId: 'visual-receipt', mode: 'save_as', overwrite: false,
        destination: { rootId: 'output', relativePath: 'old.docx' } });
      assert.equal(oldReceipt.payload.error.code, 'VALIDATION_FAILED');
      const validate = service.validateCandidate.bind(service);
      let validations = 0;
      service.validateCandidate = (...args) => { validations++; return validate(...args); };
      const published = await service.call('docx_save', save(view, 'migrated'));
      assert.equal(published.payload.data?.published, true, JSON.stringify(published));
      assert.equal(validations, 1);
      service.cleanup();
      assert(!fs.existsSync(layoutPdf));
      service.close();
      service = new OfficeService(config);
      assert.equal(service.get('meta', 'migration-3').retiredRequirements, 1, 'migration is not repeated');
      assert(fs.existsSync(path.join(s.docs, 'migrated.docx')));
    } finally { service.close(); fs.rmSync(s.root, { recursive: true, force: true }); }
  }
});

test('retiring visual policy does not bypass failed structural validation', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  let service = s.service;
  try {
    const view = await candidate(s);
    seedRetiredState(service, view, 'completed', true);
    const config = service.config;
    service.close();
    service = new OfficeService(config);
    service.validateCandidate = async () => ({ passed: false, revision: view.revision, checks: [] });
    const result = await service.call('docx_save', save(view, 'invalid'));
    assert.equal(result.payload.error.code, 'VALIDATION_FAILED');
    assert(!fs.existsSync(path.join(s.docs, 'invalid.docx')));
  } finally { service.close(); fs.rmSync(s.root, { recursive: true, force: true }); }
});

test('legacy structural validation records still publish; save uses internal validation', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    const view = await candidate(s);
    const result = await s.service.validateCandidate(s.service.doc(view.documentId), s.service.get('candidate', view.candidateId));
    s.service.put('validation', 'legacy-record', { documentId: view.documentId, candidateId: view.candidateId, revision: view.revision, result });
    const execute = s.service.execute.bind(s.service);
    s.service.execute = (name, ...args) => {
      assert.notEqual(name, 'office_validate', 'Save must use the internal validation API');
      return execute(name, ...args);
    };
    assert.equal((await s.call('docx_save', save(view, 'legacy'))).data.published, true);
    const created = await s.call('docx_create', { operationId: 'second', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'An unvalidated candidate.' }] } });
    assert.equal((await s.call('docx_save', save(created.view, 'internal'))).data.published, true);
  } finally { s.close(); }
});

test('rendering-removal migration rolls back atomically and verifies its checksum on restart', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  let service = s.service;
  try {
    const view = await candidate(s);
    seedRetiredState(service, view, 'pending');
    const config = service.config;
    service.close();
    assert.throws(() => new OfficeService(config, { checkpoint(point) {
      if (point === 'rendering-removal-before-commit') throw new Error('migration interrupted');
    } }), /migration interrupted/);
    const { DatabaseSync } = await import('./support/raw-sqlite.mjs');
    const db = new DatabaseSync(path.join(config.stateDirectory, 'state.sqlite'));
    try {
      const record = (kind, id) => JSON.parse(db.prepare('SELECT value FROM records WHERE kind=? AND id=?').get(kind, id).value);
      assert.equal(record('meta', 'schema').version, 2);
      assert.equal(record('candidate', view.candidateId).validationRequirement.validationId, 'visual-receipt');
      assert.equal(record('validation', 'visual-receipt').state, 'pending');
      assert(record('docxLayout', 'layout-old'));
    } finally { db.close(); }
    service = new OfficeService(config);
    service.put('meta', 'migration-3', { checksum: 'tampered' });
    service.close();
    assert.throws(() => new OfficeService(config), { code: 'UNSUPPORTED_STATE' });
  } finally { service.close(); fs.rmSync(s.root, { recursive: true, force: true }); }
});
