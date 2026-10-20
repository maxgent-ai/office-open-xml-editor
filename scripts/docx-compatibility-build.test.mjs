import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import test from 'node:test';
import { rolldown } from 'rolldown';
import { build } from 'vite';
import { checkDocxCompatibilityEvidence } from './check-docx-compatibility-evidence.mjs';
import {
  assertNoDocxCompatibilityDescriptions,
  docxCompatibilityMetadata,
  stripDocxCompatibilityMetadata,
} from './docx-compatibility-build.mjs';

const root = resolve(import.meta.dirname, '..');

test('bundles every source rule with the same frozen runtime ID and no metadata', async () => {
  const fixture = mkdtempSync(resolve(tmpdir(), 'docx-compatibility-build-'));
  let bundle;
  try {
    const rules = checkDocxCompatibilityEvidence(root);
    const files = [...new Set(rules.map((rule) => rule.file))];
    const entry = resolve(fixture, 'entry.mjs');
    writeFileSync(entry, files.map((file, i) =>
      `import * as rules${i} from ${JSON.stringify(resolve(root, file))};`).join('\n')
      + `\nexport const rules = [${files.map((_, i) => `rules${i}`).join(',')}]
        .flatMap(module => Object.values(module)).filter(value => value && typeof value.id === 'string');`);
    bundle = await rolldown({ input: entry, plugins: [docxCompatibilityMetadata()],
      resolve: { extensionAlias: { '.js': ['.ts', '.js'] } } });
    const { output } = await bundle.generate({ format: 'es', minify: true });
    assert.equal(output.length, 1);
    const built = resolve(fixture, 'built.mjs');
    writeFileSync(built, output[0].code);
    const actual = (await import(pathToFileURL(built).href)).rules;
    assert.deepEqual(actual.map((rule) => rule.id).sort(), rules.map((rule) => rule.id).sort());
    for (const rule of actual) {
      assert.equal(Object.isFrozen(rule), true);
      assert.deepEqual(Object.keys(rule), ['id']);
    }
  } finally {
    await bundle?.close();
    rmSync(fixture, { recursive: true, force: true });
  }
});

test('rejects shapes whose removal could hide effects or future runtime fields', () => {
  const rule = `defineCompatibilityRule({ id: 'word-test',
    evidence: { kind: 'regression-test', reference: 'path#test' }, description: 'evidence text' })`;
  for (const invalid of [
    rule.replace("description: 'evidence text'", 'description: loadDescription()'),
    rule.replace("reference: 'path#test'", 'reference: loadReference()'),
    rule.replace("id: 'word-test'", "...other, id: 'word-test'"),
    rule.replace("id: 'word-test'", "id: 'word-test', id: 'word-other'"),
    rule.replace("id: 'word-test'", "id: 'word-test', enabled: true"),
  ]) assert.throws(() => stripDocxCompatibilityMetadata(invalid, 'fixture.ts'));
});

test('bundle guard detects escaped descriptions and leaves unrelated strings alone', () => {
  const descriptions = new Set(['Evidence from Word']);
  assert.throws(() => assertNoDocxCompatibilityDescriptions(
    'export const leak = "Evidence from \\u0057ord";', descriptions, 'built.mjs'));
  assert.doesNotThrow(() => assertNoDocxCompatibilityDescriptions(
    'export const id = "word-test";', descriptions, 'built.mjs'));
});

test('transforms Windows-style module IDs including query suffixes', () => {
  const rule = checkDocxCompatibilityEvidence(root)[0];
  const file = resolve(root, rule.file);
  const code = readFileSync(file, 'utf8');
  const plugin = docxCompatibilityMetadata();
  plugin.buildStart();
  const result = plugin.transform(code, file.replaceAll('/', '\\') + '?source');
  assert.ok(result, 'Windows module must initialize and match the source authority');
  assert.doesNotThrow(() => assertNoDocxCompatibilityDescriptions(
    result.code, new Set(stripDocxCompatibilityMetadata(code, file).descriptions), file));
});

test('guards Vite worker assets emitted after the main bundle, including worker-only graphs', async () => {
  const fixture = mkdtempSync(resolve(tmpdir(), 'docx-compatibility-worker-'));
  try {
    writeFileSync(resolve(fixture, 'entry.js'),
      `export const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });`);
    writeFileSync(resolve(fixture, 'worker.js'),
      `import { WORD_SECTION_BTLR_TBRL_PAGE_FRAME as rule } from ${JSON.stringify(
        resolve(root, 'packages/docx/src/layout/compatibility.ts'))}; postMessage(rule);`);
    const options = (stripWorker) => ({
      configFile: false,
      root: fixture,
      logLevel: 'silent',
      plugins: [docxCompatibilityMetadata()],
      worker: { format: 'es', plugins: () => stripWorker ? [docxCompatibilityMetadata()] : [] },
      build: { write: false, target: 'esnext',
        lib: { entry: resolve(fixture, 'entry.js'), formats: ['es'] } },
    });
    await assert.rejects(build(options(false)), /bundled DOCX compatibility description/);
    const result = await build(options(true));
    const output = [result].flat().flatMap((bundle) => bundle.output);
    assert.ok(output.some((item) => item.type === 'asset' && /worker-.*\.js$/.test(item.fileName)),
      'the successful build must include an emitted worker asset');
  } finally {
    rmSync(fixture, { recursive: true, force: true });
  }
});

test('guards JavaScript assets emitted as bytes by a later generateBundle hook', async () => {
  const fixture = mkdtempSync(resolve(tmpdir(), 'docx-compatibility-bytes-'));
  let bundle;
  try {
    const rule = checkDocxCompatibilityEvidence(root)[0];
    const file = resolve(root, rule.file);
    const [description] = stripDocxCompatibilityMetadata(readFileSync(file, 'utf8'), file).descriptions;
    const entry = resolve(fixture, 'entry.js');
    writeFileSync(entry, 'export const value = 1;');
    bundle = await rolldown({ input: entry, plugins: [docxCompatibilityMetadata(), {
      name: 'late-byte-asset',
      generateBundle() {
        this.emitFile({ type: 'asset', fileName: 'leak.mjs',
          source: new TextEncoder().encode(`export const leak = ${JSON.stringify(description)};`) });
      },
    }] });
    await assert.rejects(bundle.generate({ format: 'es' }), /bundled DOCX compatibility description/);
  } finally {
    await bundle?.close();
    rmSync(fixture, { recursive: true, force: true });
  }
});
