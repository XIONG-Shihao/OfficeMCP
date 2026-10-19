import test from './support/test.mjs';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
import {fileURLToPath} from 'node:url';
import {nativeWorkbookBinary} from '../src/formats/xlsx/native-binary.mjs';
import {prepareEngines} from '../scripts/prepare-core.mjs';
import {xlsxSource, docxSource, nativeSource} from '../scripts/engine-source.mjs';

// OfficeMCP must run the engine that is in the source tree, never an older copy.
// These tests are the guard: if someone changes editor/xlsx or editor/docx and
// forgets `node scripts/prepare-core.mjs`, they fail and say so.
const root = fileURLToPath(new URL('../', import.meta.url));
const resolve = createRequire(import.meta.url).resolve;
const same = (a, b) => fs.readFileSync(a).equals(fs.readFileSync(b));
const hint = ' Run: node OfficeMCP/scripts/prepare-core.mjs';

test('the runtime files under .generated are exactly what the engine sources produce', () => {
  try {
    prepareEngines({ check: true });
  } catch (error) {
    assert.fail(error.message + hint);
  }
});

test('every spreadsheet engine module OfficeMCP loads is byte-identical to the configured XLSX engine core', () => {
  const names = fs.readdirSync(path.join(xlsxSource, 'core')).filter(name => name.endsWith('.mjs'));
  assert(names.length > 20, 'the engine sources were not found');
  for (const name of names)
    assert(same(resolve('#document-core/xlsx/' + name), path.join(xlsxSource, 'core', name)), name + ' differs from the source.' + hint);
});

test('the document engine OfficeMCP loads is identical to the configured DOCX engine core', () => {
  assert(same(resolve('#docx-core'), path.join(docxSource, 'core/engine.mjs')), 'engine.mjs differs from the source.' + hint);
  assert(same(resolve('#docx-observer'), path.join(docxSource, 'core/process-observer.mjs')), 'process-observer.mjs differs from the source.' + hint);
});

test('the spreadsheet engine program OfficeMCP starts is the one built from native/', () => {
  const running = nativeWorkbookBinary();
  assert(running, 'OfficeMCP has no spreadsheet engine program and would silently use the slow fallback.' + hint);
  assert(same(running, nativeSource), 'The engine program differs from the configured XLSX engine native/bin.' + hint);
});

test('there is no packed copy of an engine to go stale', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const declared = Object.keys({ ...manifest.dependencies, ...manifest.devDependencies });
  assert.deepEqual(declared.filter(name => name.startsWith('@rpf/')), []);
  assert(!fs.existsSync(path.join(root, 'vendor')), 'OfficeMCP/vendor must not come back');
  assert(!fs.existsSync(path.join(root, 'node_modules/@rpf')), 'a packed engine is installed in OfficeMCP/node_modules');
});

test('the packages the engines import are installed here at the version the editors declare', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const declaredByEditors = {
    ...JSON.parse(fs.readFileSync(path.join(xlsxSource, 'package.json'), 'utf8')).dependencies,
    ...JSON.parse(fs.readFileSync(path.join(docxSource, 'package.json'), 'utf8')).dependencies
  };
  const files = [
    ...fs.readdirSync(path.join(xlsxSource, 'core')).filter(name => name.endsWith('.mjs')).map(name => path.join(xlsxSource, 'core', name)),
    path.join(docxSource, 'core/engine.mjs'),
    path.join(docxSource, 'core/process-observer.mjs')
  ];
  const imported = new Set();
  for (const file of files)
    // An import statement is `import ... from 'name'` (the last line of it starts with `}` when it spans lines);
    // a dynamic one is `import('name')`.
    for (const match of fs.readFileSync(file, 'utf8').matchAll(/^\s*(?:import|export|\})[^\n]*?\bfrom\s*['"]([^'"\n]+)['"]|\bimport\(\s*['"]([^'"\n]+)['"]\s*\)/gm)) {
      const specifier = match[1] ?? match[2];
      if (/^(node:|[./#])/.test(specifier)) continue;
      const parts = specifier.split('/');
      imported.add(specifier.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0]);
    }
  assert(imported.has('fflate') && imported.has('@xmldom/xmldom'), 'the import scan found nothing');
  for (const name of imported) {
    assert(manifest.dependencies[name], `${name} is imported by an engine but OfficeMCP does not depend on it`);
    assert.equal(manifest.dependencies[name], declaredByEditors[name], `${name}: OfficeMCP and the editor declare different versions`);
  }
});
