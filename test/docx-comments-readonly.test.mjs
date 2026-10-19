import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { Paragraph } from 'docx';
import { createServiceFixture } from './service-fixture.mjs';
import { commentedDocument } from './comment-fixture.mjs';
import { DOCX_AGENT_ACTIONS, docxActionDiscovery } from '../src/formats/docx/docx-agent-actions.mjs';
import { manifest } from '../src/contracts/registry.mjs';

test('comment writes are absent from discovery and rejected before a standalone session can mutate', async () => {
  const writers = ['add_comments', 'resolve_comments', 'reply_to_comment', 'comment_paragraphs', 'create_comment', 'edit_comment', 'delete_comment', 'update_comment', 'resolve_comment', 'reopen_comment'];
  assert.equal(DOCX_AGENT_ACTIONS.length, 79);
  for (const name of writers) {
    assert(!DOCX_AGENT_ACTIONS.includes(name));
    assert.throws(() => docxActionDiscovery({ actions: [name] }), /Unknown/);
  }
  for (const preset of ['all', 'full', 'docx-agent']) {
    const tools = manifest(preset);
    assert(!tools.some(t => /comment.*(create|edit|delete|resolve|reply)|(?:create|edit|delete|resolve|reply).*comment/.test(t.name)));
    if (preset === 'all') continue;
    const f = createServiceFixture({ toolset: preset });
    try {
      const file = path.join(f.docs, 'existing.docx'), bytes = await commentedDocument({ after: [new Paragraph('Unrelated original text.')] });
      fs.writeFileSync(file, bytes); const opened = await f.call('docx_open', { path: file });
      const read = await f.call('docx_read', { view: opened.view, representation: 'domain', domain: 'comments', limit: 20 });
      assert(read.data.items.some(c => c.text === 'Keep this actual reviewer note.'));
      const query = await f.call('docx_query', { view: opened.view, select: { type: 'entity', domain: 'comments', text: 'Keep this actual reviewer note.' }, require: 'exactlyOne', limit: 2 });
      assert.equal(query.data.items[0].text, 'Keep this actual reviewer note.');
      for (const action of writers) {
        const result = await f.service.call('docx_apply', { view: opened.view, operationId: action, action, arguments: {} });
        assert(result.isError); assert.equal(result.payload.error.code, 'INVALID_ARGUMENT');
      }
      const batch = await f.service.call('docx_apply', { view: opened.view, operationId: 'batch', steps: [
        { action: 'insert_heading', arguments: { text: 'Must not appear', level: 1 } },
        { action: 'add_comments', arguments: { commentText: 'Forbidden' } },
      ] });
      assert(batch.isError); assert.equal(f.service.rows('candidate').length, 0); assert.deepEqual(fs.readFileSync(file), bytes);
      if (preset === 'full') {
      const legacy = await f.call('docx_read_details', { view: opened.view, domain: 'comments', limit: 20 });
      assert(legacy.data.items.some(c => c.text === 'Keep this actual reviewer note.'));
      }
    } finally { f.close(); }
  }
});

test('editing unrelated text and saving preserves an existing authored comment and its open status', async () => {
  const f = createServiceFixture({ toolset: 'docx-agent' });
  try {
    const file = path.join(f.docs, 'existing.docx'), bytes = await commentedDocument({ after: [new Paragraph('Unrelated original text.')] });
    fs.writeFileSync(file, bytes); let view = (await f.call('docx_open', { path: file })).view;
    view = (await f.call('docx_apply', { view, operationId: 'edit', action: 'replace_text', select: { type: 'text', pattern: 'Unrelated original text.' },
      arguments: { edits: [{ find: 'original', replace: 'updated' }], caseSensitive: true, changeMode: 'direct' } })).view;
    const saved = await f.call('docx_save', { view, operationId: 'save', destination: { rootId: 'output', relativePath: 'saved.docx' } });
    const reopened = (await f.call('docx_open', { path: saved.data.path })).view;
    const comments = await f.call('docx_read', { view: reopened, representation: 'domain', domain: 'comments', limit: 20 });
    const comment = comments.data.items.find(c => c.text === 'Keep this actual reviewer note.'); assert(comment); assert.equal(comment.status, 'open');
    assert.deepEqual(fs.readFileSync(file), bytes);
  } finally { f.close(); }
});

test('generic and legacy readers retain existing-comment access', async () => {
  for (const toolset of ['all', 'recommended']) {
    const f = createServiceFixture({ toolset });
    try {
      const file = path.join(f.docs, 'existing.docx'); fs.writeFileSync(file, await commentedDocument());
      const opened = await f.call('office_open', { path: file });
      const read = await f.call('office_read', { documentId: opened.document.id, docxDomain: 'comments', limit: 20 });
      assert(read.data.items.some(c => c.text === 'Keep this actual reviewer note.'));
      if (toolset === 'all') {
        const details = await f.call('docx_read_details', { view: read.view, domain: 'comments', limit: 20 });
        assert(details.data.items.some(c => c.text === 'Keep this actual reviewer note.'));
      }
    } finally { f.close(); }
  }
});
