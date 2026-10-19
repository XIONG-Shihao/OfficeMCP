import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { saveChecked } from '../src/formats/docx/docx-sdk.mjs';
import { createServiceFixture } from './service-fixture.mjs';
import { unzipSync, strFromU8 } from 'fflate';
import { Paragraph } from 'docx';
import { commentedDocument } from './comment-fixture.mjs';

const story = { kind: 'story', storyType: 'body' };
const change = { id: 'change-1', type: 'insertion', revisionGroupId: 'word-insertion-1', author: 'Reviewer', date: '2026-09-22T00:00:00.000Z', insertedText: ' appended', navigationTarget: { story } };
const projection = {
  commentId: 'tc-comment:word-insertion-1', status: 'open', trackedChange: true,
  trackedChangeParentId: change.id, creatorName: change.author, createdTime: Date.parse(change.date),
  trackedChangeType: 'insert', trackedChangeSide: 'inserted', trackedChangeStory: story,
  trackedChangeAnchorKey: change.revisionGroupId, trackedChangeText: change.insertedText,
  insertedText: change.insertedText, anchoredText: change.insertedText,
  trackedChangeLink: { trackedChange: true, trackedChangeId: change.id, trackedChangeType: change.type,
    side: 'inserted', trackedChangeStory: story, trackedChangeAnchorKey: change.revisionGroupId,
    insertedText: change.insertedText, trackedChangeText: change.insertedText },
};
const realComment = { commentId: 'authored-1', text: 'Keep this actual reviewer note.', creatorName: 'Reviewer' };
function fixture({ before = [realComment, projection], after = [realComment], afterChange = change, afterText = 'Body appended' } = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'docx-tracked-snapshot-'));
  const task = { input: path.join(root, 'input.docx'), output: path.join(root, 'output.docx'), maxFileBytes: 4096, args: { steps: [] } };
  fs.writeFileSync(task.input, 'source');
  const document = (comments, tracked, text) => {
    const doc = { get: async () => ({ body: [{ text }] }), close: async () => {},
      save: async () => fs.writeFileSync(task.output, 'candidate'),
      headerFooters: { parts: { list: async () => ({ items: [], total: 0 }) } } };
    for (const name of ['images', 'sections', 'footnotes', 'fields', 'hyperlinks', 'contentControls', 'trackChanges', 'comments', 'captions', 'crossRefs']) {
      const items = name === 'trackChanges' ? [tracked] : name === 'comments' ? comments : [];
      doc[name] = { list: async () => ({ items, total: items.length }) };
    }
    return doc;
  };
  return { run: () => saveChecked({ open: async () => document(after, afterChange, afterText) }, document(before, change, 'Body appended'), task, 'source-revision'), close: () => fs.rmSync(root, { recursive: true, force: true }) };
}

test('only a verified SDK tracked-comment projection may disappear on reopen', async () => {
  const f = fixture();
  try { assert((await f.run()).revision); } finally { f.close(); }
});

test('authored comments, forged projections, tracked content and body mismatches still fail verification', async () => {
  const cases = [
    { after: [] },
    { after: [{ ...realComment, text: 'Lost reviewer note' }] },
    { before: [realComment, { ...projection, text: 'Real note with a misleading ID' }] },
    { before: [realComment, { ...projection, insertedText: 'forged content' }] },
    { before: [realComment, { ...projection, creatorName: 'Wrong author' }] },
    { before: [realComment, { ...projection, trackedChangeParentId: 'unrelated-change' }] },
    { before: [realComment, { ...projection, trackedChangeLink: undefined }] },
    { before: [realComment, { ...projection, trackedChangeStory: undefined }] },
    { afterChange: { ...change, insertedText: 'wrong tracked content' } },
    { afterText: 'Wrong saved body' },
  ];
  for (const options of cases) {
    const f = fixture(options);
    try { await assert.rejects(f.run(), { code: 'VALIDATION_FAILED' }); } finally { f.close(); }
  }
});

test('tracked suffix insertion preserves authored comments and body after real SDK save/reopen', async () => {
  const f = createServiceFixture({ toolset: 'full' });
  try {
    const source = path.join(f.docs, 'existing-review-note.docx');
    fs.writeFileSync(source, await commentedDocument({ after: [new Paragraph('Track this original clause.')] }));
    let view = (await f.call('docx_open', { path: source })).view;
    view = (await f.call('docx_apply', { view, operationId: 'tracked-suffix', action: 'replace_text',
      select: { type: 'text', pattern: 'Track this original clause.' }, author: 'Reviewer',
      arguments: { edits: [{ find: 'Track this original clause.', replace: 'Track this original clause. Appended review text.' }], changeMode: 'tracked' } })).view;
    assert.equal((await f.call('docx_validate', { view })).data.passed, true);
    const saved = await f.call('docx_save', { view, operationId: 'save', destination: { rootId: 'output', relativePath: 'reviewed.docx' } });
    const parts = unzipSync(fs.readFileSync(saved.data.path));
    assert.match(strFromU8(parts['word/document.xml']), /<w:ins\b/);
    assert(strFromU8(parts['word/document.xml']).includes('Appended review text.'));
    assert(strFromU8(parts['word/comments.xml']).includes('Keep this actual reviewer note.'));
    await f.call('docx_close', { view, disposition: 'retain' });
    view = (await f.call('docx_open', { path: saved.data.path })).view;
    const comments = await f.call('docx_read', { view, representation: 'domain', domain: 'comments' });
    assert(comments.data.items.some(item => item.text === 'Keep this actual reviewer note.'));
    assert((await f.call('docx_read', { view, representation: 'domain', domain: 'tracked_changes' })).data.total > 0);
    assert((await f.call('docx_read', { view, representation: 'text' })).data.items.some(item => item.text.includes('Appended review text.')));
  } finally { f.close(); }
});
