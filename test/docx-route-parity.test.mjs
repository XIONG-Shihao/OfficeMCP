import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createServiceFixture } from './service-fixture.mjs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { Document, Packer, Paragraph, TextRun, Footer, Table, TableRow, TableCell } from 'docx';
import { readDocxDomain } from '../src/formats/docx/docx-domain-items.mjs';
import { queryAgentDocument } from '../src/formats/docx/docx-agent-query.mjs';

for (const placement of ['before', 'after'])
  test(`legacy and agent paragraph insertion preserve ${placement} semantics through save and reopen`, async () => {
    const s = createServiceFixture({ toolset: 'full' });
    try {
      const seed = await s.call('docx_create', { operationId: 'seed', document: { format: 'docx', blocks: [
        { kind: 'paragraph', text: 'Anchor paragraph.' }, { kind: 'paragraph', text: 'Untouched paragraph.' },
      ] } });
      const saved = await s.call('docx_save', { view: seed.view, operationId: 'seed-save', destination: { rootId: 'output', relativePath: 'seed.docx' } });
      const original = fs.readFileSync(saved.data.path);
      const texts = [];
      for (const route of ['legacy', 'agent']) {
        const opened = await s.call('docx_open', { path: saved.data.path });
        let view = opened.view;
        if (route === 'legacy') {
          const args = { documentId: view.documentId, expectedSourceRevision: view.revision, operationId: 'legacy-edit', paragraphIndex: 0,
            expectedText: 'Anchor paragraph.', text: 'Inserted paragraph.', placement, preservationProfile: 'docx.superdoc.v1' };
          const refused = await s.service.call('docx_insert_paragraph', { ...args, operationId: 'wrong-anchor', expectedText: 'Wrong text.' });
          assert.equal(refused.payload.error.code, 'TARGET_CONFLICT');
          const edited = await s.call('docx_insert_paragraph', args);
          view = { documentId: view.documentId, candidateId: edited.data.candidateId, revision: edited.data.candidateRevision };
          assert.deepEqual((await s.call('docx_insert_paragraph', args)).data, edited.data);
        } else {
          const query = await s.call('docx_query', { view, select: { type: 'text', pattern: 'Anchor paragraph.' }, require: 'exactlyOne' });
          view = (await s.call('docx_apply', { view, operationId: 'agent-edit', action: 'insert_paragraphs', selectorTarget: query.data.items[0].target,
            arguments: { text: 'Inserted paragraph.', placement: { at: placement } } })).view;
        }
        const output = await s.call('docx_save', { view, operationId: `${route}-save`, destination: { rootId: 'output', relativePath: `${route}.docx` } });
        const reopened = await s.call('docx_open', { path: output.data.path });
        const read = await s.call('docx_read', { view: reopened.view, representation: 'linear' });
        texts.push(read.data.items.map(item => item.text).filter(Boolean));
      }
      assert.deepEqual(texts[0], texts[1]);
      assert.deepEqual(texts[0], placement === 'before'
        ? ['Inserted paragraph.', 'Anchor paragraph.', 'Untouched paragraph.']
        : ['Anchor paragraph.', 'Inserted paragraph.', 'Untouched paragraph.']);
      assert.deepEqual(fs.readFileSync(saved.data.path), original);
    } finally { s.close(); }
  });

import { docxWorker } from '../src/formats/docx/docx-bridge.mjs';
test('legacy table edits resolve tables without scanning unrelated paragraphs', async () => {
  const domains = [];
  const service = { docxPool: { run: async (kind, _input, args) => {
    if (kind === 'docx_native_read') {
      domains.push(args.domain);
      return { items: [{ nativeTarget: { kind: 'table', address: { nodeId: 'table' } } }] };
    }
    return args.steps;
  } } };
  const steps = await docxWorker(service, 'docx_table_cell', '/fixture.docx', {
    target: { tableIndex: 0, row: 0, column: 0 }, expectedText: '125', replacement: '150',
  }, {});
  assert.deepEqual(domains, ['tables']);
  assert.deepEqual(steps[0].cells, [{ row: 0, column: 0, expectedText: '125', text: '150' }]);
});

test('native and agent routes preserve formatted replacements, table guards, review decisions and paged reads', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    const seed = path.join(s.docs, 'rich-route-source.docx');
    fs.writeFileSync(seed, await Packer.toBuffer(new Document({ sections: [{
      footers: { default: new Footer({ children: [new Paragraph('Shared footer')] }) },
      children: [
        new Paragraph({ children: [new TextRun('Amount '), new TextRun({ text: '100', bold: true }), new TextRun({ text: ' USD.', italics: true })] }),
        new Paragraph('Original clause.'),
        new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph('Service')] }), new TableCell({ children: [new Paragraph('100')] })] })] }),
      ],
    }] })));
    const original = fs.readFileSync(seed);
    const results = [];
    for (const route of ['native', 'agent']) {
      let view = (await s.call('docx_open', { path: seed })).view;
      const replace = async (expectedText, replacement, paragraphIndex, tracked, operationId) => {
        if (route === 'native') return (await s.call('office_edit', { view, operationId, action: {
          type: 'replace_text', target: { partName: 'word/document.xml', paragraphIndex }, expectedText, replacement, tracked,
        } })).view;
        const target = (await s.call('docx_query', { view, select: { type: 'text', pattern: expectedText }, require: 'exactlyOne' })).data.items[0].target;
        return (await s.call('docx_apply', { view, operationId, action: 'replace_text', selectorTarget: target,
          arguments: { edits: [{ find: expectedText, replace: replacement }], caseSensitive: true, changeMode: tracked ? 'tracked' : 'direct' } })).view;
      };
      view = await replace('Amount 100 USD.', 'Amount 125 USD.', 0, false, `${route}-format`);
      const wrong = route === 'native'
        ? await s.service.call('office_edit', { view, operationId: `${route}-wrong-table`, action: { type: 'table_cell', target: { partName: 'word/document.xml', tableIndex: 0, row: 0, column: 1 }, expectedText: 'wrong', text: '150' } })
        : await s.service.call('docx_apply', { view, operationId: `${route}-wrong-table`, action: 'write_table_cells', select: { type: 'table', header: 'Service' }, arguments: { cells: [{ row: 0, column: 1, expectedText: 'wrong', text: '150' }] } });
      assert.equal(wrong.payload.error.code, 'TARGET_CONFLICT');
      if (route === 'native') view = (await s.call('office_edit', { view, operationId: `${route}-table`, action: { type: 'table_cell', target: { partName: 'word/document.xml', tableIndex: 0, row: 0, column: 1 }, expectedText: '100', text: '150' } })).view;
      else view = (await s.call('docx_apply', { view, operationId: `${route}-table`, action: 'write_table_cells', select: { type: 'table', header: 'Service' }, arguments: { cells: [{ row: 0, column: 1, expectedText: '100', text: '150' }] } })).view;
      view = await replace('Original clause.', 'Revised clause.', 1, true, `${route}-tracked`);
      if (route === 'native') {
        const changes = await s.call('docx_read_details', { view, domain: 'tracked_changes' });
        const changeIds = [...new Set(changes.data.items.flatMap(item => Object.values(item.wordRevisionIds ?? {})).filter(Boolean))];
        assert(changeIds.length > 0);
        view = (await s.call('office_edit', { view, operationId: `${route}-review`, action: { type: 'review_changes', partName: 'word/document.xml', changeIds, decision: 'accept' } })).view;
      } else {
        const changes = await s.call('docx_read', { view, representation: 'domain', domain: 'tracked_changes' });
        assert.equal(changes.data.items.length, 1);
        view = (await s.call('docx_apply', { view, operationId: `${route}-review`, action: 'accept_tracked_change', selectorTarget: changes.data.items[0].target, arguments: {} })).view;
      }
      const saved = await s.call('docx_save', { view, operationId: `${route}-rich-save`, destination: { rootId: 'output', relativePath: `${route}-rich.docx` } });
      view = (await s.call('docx_open', { path: saved.data.path })).view;
      const formatted = await s.call('docx_query', { view, select: { type: 'text', pattern: '125' }, require: 'exactlyOne' });
      assert.equal(formatted.data.items[0].text, '125');
      const suffix = await s.call('docx_query', { view, select: { type: 'text', pattern: ' USD.' }, require: 'exactlyOne' });
      assert.equal(suffix.data.items[0].text, ' USD.');
      const xml = new DOMParser().parseFromString(strFromU8(unzipSync(fs.readFileSync(saved.data.path))['word/document.xml']), 'application/xml');
      const runs = Array.from(xml.getElementsByTagName('w:r'));
      const textOfRun = run => Array.from(run.getElementsByTagName('w:t')).map(node => node.textContent).join('');
      assert(runs.some(run => textOfRun(run) === '125' && run.getElementsByTagName('w:b').length));
      assert(runs.some(run => textOfRun(run) === ' USD.' && run.getElementsByTagName('w:i').length));
      const nativePage = await s.call('docx_structure', { documentId: view.documentId, expectedRevision: view.revision, offset: 1, limit: 1 });
      const agentPage = await s.call('docx_read', { view, representation: 'domain', domain: 'paragraphs', offset: 1, limit: 1 });
      assert.deepEqual(nativePage.data.nodes.map(item => item.text), agentPage.data.items.map(item => item.text));
      assert.equal(nativePage.data.nextOffset, agentPage.data.nextOffset);
      const story = await s.call('docx_read_details', { view, domain: 'headers_footers' });
      assert(story.data.items.some(item => item.text === 'Shared footer'));
      const body = await s.call('docx_read', { view, representation: 'domain', domain: 'paragraphs' });
      results.push(body.data.items.map(item => item.text));
    }
    assert.deepEqual(results[0], results[1]);
    assert.deepEqual(results[0], ['Amount 125 USD.', 'Revised clause.', 'Service', '150']);
    assert.deepEqual(fs.readFileSync(seed), original);
  } finally { s.close(); }
});

test('domain reads and entity queries reuse the same paginated mutation locators', async () => {
  const calls = [];
  const items = [
    { id: 'comment-a', text: 'First', author: 'Editor' },
    { id: 'comment-b', text: 'Second', author: 'Editor' },
  ];
  const doc = { comments: { list: async ({ offset }) => { calls.push(offset); return { items: items.slice(offset, offset + 1), total: items.length }; } } };
  const domain = await readDocxDomain(doc, 'comments');
  const queried = await queryAgentDocument(doc, { select: { type: 'entity', domain: 'comments', author: 'Editor', caseSensitive: true }, require: 'all', offset: 1, limit: 1 });
  assert.deepEqual(queried.items[0].nativeTarget, domain[1].nativeTarget);
  assert.equal(queried.items[0].text, 'Second');
  assert.equal(queried.total, 2);
  assert.equal(queried.nextOffset, null);
  assert.deepEqual(calls, [0, 1, 0, 1]);
});

test('native and agent failed plans discard an earlier successful mutation before actor reuse', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    const seed = await s.call('docx_create', { operationId: 'failure-seed', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'Alpha Beta' }] } });
    const saved = await s.call('docx_save', { view: seed.view, operationId: 'failure-seed-save', destination: { rootId: 'output', relativePath: 'failure-source.docx' } });
    const source = fs.readFileSync(saved.data.path);
    for (const route of ['native', 'agent']) {
      const view = (await s.call('docx_open', { path: saved.data.path })).view;
      if (route === 'native') {
        const current = s.service.view({ documentId: view.documentId, expectedRevision: view.revision });
        const inventory = await s.service.worker('docx_native_read', current.file, { documentId: view.documentId, domain: 'paragraphs', offset: 0, limit: 1 }, {});
        const target = inventory.items[0].nativeTarget;
        const output = path.join(s.root, 'failed-native-output.docx');
        await assert.rejects(s.service.worker('docx_native_edit', current.file, {
          documentId: view.documentId,
          steps: [
            { type: 'text', action: 'replace', expectedText: 'Beta', content: 'Gamma', changeMode: 'direct', target: { id: 'first' }, locator: { ...target, kind: 'text', start: 6, end: 10, text: 'Beta' } },
            { type: 'text', action: 'replace', expectedText: 'Alpha Beta', content: 'Overlapping', changeMode: 'direct', target: { id: 'second' }, locator: { ...target, kind: 'text', start: 0, end: 10, text: 'Alpha Beta' } },
          ],
        }, { output }), { code: 'TARGET_CONFLICT' });
        assert.equal(fs.existsSync(output), false);
      } else {
        const failed = await s.service.call('docx_apply', { view, operationId: 'failed-agent-plan', steps: [
          { action: 'replace_text', select: { type: 'text', pattern: 'Beta' }, arguments: { edits: [{ find: 'Beta', replace: 'Gamma' }], caseSensitive: true } },
          { action: 'replace_text', select: { type: 'text', pattern: 'missing' }, arguments: { edits: [{ find: 'missing', replace: 'Never' }], caseSensitive: true } },
        ] });
        assert(failed.isError);
        assert(['TARGET_CONFLICT', 'ENGINE_FAILED'].includes(failed.payload.error.code));
      }
      const after = await s.call('docx_read', { view, representation: 'domain', domain: 'paragraphs' });
      assert.deepEqual(after.data.items.map(item => item.text), ['Alpha Beta']);
      assert.deepEqual(fs.readFileSync(saved.data.path), source);
    }
  } finally { s.close(); }
});

test('delete_text and insert_heading preserve their dedicated action behavior', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    let view = (await s.call('docx_create', { operationId: 'alias-seed', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'Keep REMOVE Keep' }] } })).view;
    const removed = await s.call('docx_apply', { view, operationId: 'alias-delete', action: 'delete_text', select: { type: 'text', pattern: 'REMOVE' }, arguments: { finds: ['REMOVE'], caseSensitive: true } });
    assert.equal(removed.data.action, 'delete_text');
    view = removed.view;
    const heading = await s.call('docx_apply', { view, operationId: 'alias-heading', action: 'insert_heading', select: { type: 'text', pattern: 'Keep  Keep' }, arguments: { text: 'Added heading', level: 2, placement: { at: 'after' } } });
    assert.equal(heading.data.action, 'insert_heading');
    const read = await s.call('docx_read', { view: heading.view, representation: 'domain', domain: 'paragraphs' });
    assert.deepEqual(read.data.items.map(item => item.text), ['Keep  Keep', 'Added heading']);
    assert.equal(read.data.items[1].kind, 'heading');
  } finally { s.close(); }
});

test('insert_heading adds no paragraph spacer in blank or one-paragraph documents', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    for (const [seedIndex, text] of ['', 'Existing paragraph'].entries()) {
      const created = await s.call('docx_create', { operationId: `heading-source-${seedIndex}`, document: { format: 'docx', blocks: [{ kind: 'paragraph', text }] } });
      const saved = await s.call('docx_save', { view: created.view, operationId: `heading-source-save-${seedIndex}`, destination: { rootId: 'output', relativePath: `heading-source-${seedIndex}.docx` } });
      const sourceView = (await s.call('docx_open', { path: saved.data.path })).view;
      const before = await s.call('docx_read', { view: sourceView, representation: 'domain', domain: 'paragraphs' });
      assert.equal(before.data.items.length, 1);
      for (const [placementIndex, placement] of [undefined, { at: 'document_start' }, { at: 'document_end' }].entries()) {
        const inserted = await s.call('docx_apply', { view: sourceView, operationId: `heading-${seedIndex}-${placementIndex}`, action: 'insert_heading', arguments: { text: 'Only heading', level: 2, ...(placement ? { placement } : {}) } });
        const after = await s.call('docx_read', { view: inserted.view, representation: 'domain', domain: 'paragraphs' });
        assert.equal(after.data.items.filter(item => item.kind === 'heading' && item.text === 'Only heading').length, 1);
        assert.equal(after.data.items.length, 2, 'A heading adds one block, without the insert_paragraphs drafting spacer');
        assert.equal(after.data.items.filter(item => item.text === '').length, text === '' ? 1 : 0);
      }
    }
  } finally { s.close(); }
});

test('delete_text removes every occurrence in the selected block and retains its whitespace guard', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    let view = (await s.call('docx_create', { operationId: 'delete-repeat-source', document: { format: 'docx', blocks: [{ kind: 'paragraph', text: 'X X' }] } })).view;
    const refused = await s.service.call('docx_apply', { view, operationId: 'unscoped-whitespace-delete', action: 'delete_text', arguments: { finds: [' '] } });
    assert(refused.isError, 'Unscoped whitespace deletion must not become a general replace_text action');
    const target = (await s.call('docx_query', { view, select: { type: 'node', nodeType: 'paragraph' }, require: 'exactlyOne' })).data.items[0].target;
    view = (await s.call('docx_apply', { view, operationId: 'delete-repeat', action: 'delete_text', selectorTarget: target, arguments: { finds: ['X'], caseSensitive: true } })).view;
    const after = await s.call('docx_read', { view, representation: 'domain', domain: 'paragraphs' });
    assert.deepEqual(after.data.items.map(item => item.text), [' ']);
  } finally { s.close(); }
});

test('native and agent story replacements edit the selected footer without matching body text', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    const seed = await s.call('docx_create', { operationId: 'story-seed', document: { format: 'docx', footer: 'Shared footer', blocks: [{ kind: 'paragraph', text: 'Shared footer' }] } });
    const saved = await s.call('docx_save', { view: seed.view, operationId: 'story-seed-save', destination: { rootId: 'output', relativePath: 'story-source.docx' } });
    const source = fs.readFileSync(saved.data.path);
    for (const route of ['native', 'agent']) {
      let view = (await s.call('docx_open', { path: saved.data.path })).view;
      const output = path.join(s.docs, `${route}-story.docx`);
      if (route === 'native') {
        const current = s.service.view({ documentId: view.documentId, expectedRevision: view.revision });
        const inventory = await s.service.worker('docx_native_read', current.file, { documentId: view.documentId, domain: 'headers_footers', offset: 0, limit: 30 }, {});
        const footer = inventory.items.find(item => item.kind === 'footer' && item.nativeTarget);
        assert(footer);
        await s.service.worker('docx_native_edit', current.file, { documentId: view.documentId, steps: [{
          type: 'headers_footers', expectedText: 'Shared footer', text: 'Updated footer', target: { id: 'footer' }, locator: footer.nativeTarget,
        }] }, { output });
      } else {
        const section = (await s.call('docx_read', { view, representation: 'domain', domain: 'sections' })).data.items[0];
        view = (await s.call('docx_apply', { view, operationId: 'agent-story', action: 'replace_header_footer_text', arguments: {
          sectionId: section.itemId, kind: 'footer', find: 'Shared footer', replace: 'Updated footer',
        } })).view;
        await s.call('docx_save', { view, operationId: 'agent-story-save', destination: { rootId: 'output', relativePath: 'agent-story.docx' } });
      }
      view = (await s.call('docx_open', { path: output })).view;
      const body = await s.call('docx_read', { view, representation: 'domain', domain: 'paragraphs' });
      assert.equal(body.data.items[0].text, 'Shared footer');
      const stories = await s.call('docx_read', { view, representation: 'domain', domain: 'headers_footers' });
      assert(stories.data.items.some(item => item.text === 'Updated footer'));
      assert.deepEqual(fs.readFileSync(saved.data.path), source);
    }
  } finally { s.close(); }
});

test('native and agent table row deletion and insertion have the same operation semantics', async () => {
  const s = createServiceFixture({ toolset: 'full' });
  try {
    const seed = await s.call('docx_create', { operationId: 'rows-seed', document: { format: 'docx', blocks: [{ kind: 'table', rows: [['Keep', '1'], ['Delete', '2'], ['Stay', '3']] }] } });
    const saved = await s.call('docx_save', { view: seed.view, operationId: 'rows-save', destination: { rootId: 'output', relativePath: 'rows.docx' } });
    for (const route of ['native', 'agent']) {
      let view = (await s.call('docx_open', { path: saved.data.path })).view;
      if (route === 'native') view = (await s.call('office_edit', { view, operationId: 'native-delete-row', action: { type: 'table_row', target: { partName: 'word/document.xml', tableIndex: 0, row: 1 }, action: 'delete' } })).view;
      else view = (await s.call('docx_apply', { view, operationId: 'agent-delete-row', action: 'delete_table_row', select: { type: 'table', contains: 'Delete' }, arguments: { rowIndex: 1 } })).view;
      let tables = await s.call('docx_read', { view, representation: 'domain', domain: 'tables' });
      assert.deepEqual(tables.data.items[0].rows.map(row => row.map(cell => cell.text)), [['Keep', '1'], ['Stay', '3']]);
      if (route === 'native') view = (await s.call('office_edit', { view, operationId: 'native-insert-row', action: { type: 'table_row', target: { partName: 'word/document.xml', tableIndex: 0, row: 1 }, action: 'insert_before', values: ['New', '4'] } })).view;
      else view = (await s.call('docx_apply', { view, operationId: 'agent-insert-row', action: 'insert_table_row', select: { type: 'table', contains: 'Stay' }, arguments: { rowIndex: 1, position: 'before', cellTexts: ['New', '4'] } })).view;
      tables = await s.call('docx_read', { view, representation: 'domain', domain: 'tables' });
      assert.deepEqual(tables.data.items[0].rows.map(row => row.map(cell => cell.text)), [['Keep', '1'], ['New', '4'], ['Stay', '3']]);
      await s.call('docx_save', { view, operationId: `${route}-rows-final`, destination: { rootId: 'output', relativePath: `${route}-rows-final.docx` } });
    }
  } finally { s.close(); }
});
