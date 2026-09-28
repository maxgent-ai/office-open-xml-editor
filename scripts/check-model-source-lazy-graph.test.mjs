import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { build } from 'vite';
import { assertDispatchCost, assertLazySourceOwner, assertNoTopLevelModelImport, assertEagerSourceEntries, assertNoSourceRuntime, checkSourceDispatchImports } from './check-model-source-lazy-graph.mjs';

const tempRoot = process.platform === 'darwin' ? '/private/tmp' : tmpdir();

test('a static owner import in the worker graph fails, while a selected-source import stays lazy', () => {
  const root = mkdtempSync(join(tempRoot, 'ooxml-source-graph-'));
  try {
    const internal = join(root, 'internal');
    mkdirSync(internal);
    const owner = join(internal, 'worker-document-source.ts');
    const worker = join(root, 'worker.ts');
    writeFileSync(owner, 'export class Owner {}\n');
    writeFileSync(worker, "if (source) await import('./internal/worker-document-source.js');\n");
    assert.doesNotThrow(() => assertLazySourceOwner(worker, owner));
    writeFileSync(worker, "import { Owner } from './internal/worker-document-source.js';\nif (source) await import('./internal/worker-document-source.js');\n");
    assert.throws(() => assertLazySourceOwner(worker, owner), /statically reaches/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('named source dispatch allowlist rejects imports outside the owner', () => {
  const path = 'packages/docx/src/document.ts';
  const dispatch = "class DocxDocument { static async load(opts) { if (opts.modelSources) return import('./internal/document-model-source.js'); } }";
  assert.doesNotThrow(() => checkSourceDispatchImports([{ path, text: dispatch }]));
  assert.throws(() => checkSourceDispatchImports([{
    path, text: `${dispatch}\nvoid import('./internal/document-model-source.js');`,
  }]), /outside an allowed dispatch function/);
  assert.doesNotThrow(() => checkSourceDispatchImports([{
    path, text: "class DocxDocument { static async load(opts) { return import('./internal/document-model-source.js'); } }",
  }]));
});

test('rejects an unconditional model-source import in a built entry', () => {
  assert.throws(() => assertNoTopLevelModelImport(
    "void import('./document-model-source-CYLDKzzV.js');", 'dist/docx.mjs',
  ), /top level/);
  assert.doesNotThrow(() => assertNoTopLevelModelImport(
    "async function load(opts) { if (opts.modelSources) await import('./document-model-source-CYLDKzzV.js'); }",
    'dist/docx.mjs',
  ));
});

test('each eager entry type rejects a transitive static model-source import', () => {
  const root = mkdtempSync(join(tempRoot, 'ooxml-source-entry-'));
  try {
    const owner = join(root, 'internal', 'document-model-source.ts');
    mkdirSync(join(root, 'internal'));
    writeFileSync(owner, 'export const source = true;\n');
    for (const name of ['core-index', 'package-index', 'worker', 'render-worker', 'node-index']) {
      const entry = join(root, `${name}.ts`);
      const bridge = join(root, `${name}-bridge.ts`);
      writeFileSync(entry, `import './${name}-bridge.js';\n`);
      writeFileSync(bridge, "import './internal/document-model-source.js';\n");
      assert.throws(() => assertEagerSourceEntries([entry]), /statically reaches/);
      writeFileSync(bridge, "void import('./internal/document-model-source.js');\n");
      assert.doesNotThrow(() => assertEagerSourceEntries([entry]));
    }
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('a model-source marker leaking into eager emitted code fails', () => {
  assert.throws(() => assertNoSourceRuntime('const marker="ooxml-model-source-module/v1";', 'entry'), /optional source runtime/);
});

test('a model-source helper factored outside the branch is charged to the build difference', async () => {
  const root = mkdtempSync(join(tempRoot, 'ooxml-source-cost-'));
  try {
    const entry = join(root, 'entry.js');
    const enabledDir = join(root, 'enabled');
    const disabledDir = join(root, 'disabled');
    const lookup = Array.from({ length: 300 }, (_, index) =>
      `k${index}: 'value-${index}-${(index * 7919).toString(36)}'`).join(',');
    writeFileSync(entry, `function prepareSource(key) { return ({${lookup}})[key]; }
export async function load(options = {}) {
  if (__OOXML_MODEL_SOURCES__ && options.modelSources !== undefined) {
    prepareSource(options.key);
    return import('./model-source.js');
  }
  return options.key;
}
`);
    writeFileSync(join(root, 'model-source.js'), 'export const source = true;\n');
    for (const [flag, outDir] of [['true', enabledDir], ['false', disabledDir]]) {
      await build({ configFile: false, logLevel: 'silent', define: { __OOXML_MODEL_SOURCES__: flag },
        build: { outDir, lib: { entry, formats: ['es'], fileName: () => 'entry.mjs' } } });
    }
    assert.throws(() => assertDispatchCost(join(enabledDir, 'entry.mjs'),
      join(disabledDir, 'entry.mjs'), 3_000, 'fixture'), /dispatch cost exceeds/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
