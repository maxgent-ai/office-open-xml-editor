import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { build } from 'vite';
import { moduleStatsPlugin } from './module-stats-plugin.mjs';
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
  const dispatch = "class DocxDocument { static async load(opts) { if (opts.legacyDoc) return import('./internal/document-model-source.js'); } }";
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
    "async function load(opts) { if (opts.legacyDoc) await import('./document-model-source-CYLDKzzV.js'); }",
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
  assert.throws(() => assertNoSourceRuntime(JSON.stringify('ooxml-model-source-module/v2'), 'entry'), /optional source runtime/);
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
  if (__OOXML_MODEL_SOURCES__ && options.legacyDoc !== undefined) {
    prepareSource(options.key);
    return import('./model-source.js');
  }
  return options.key;
}
`);
    writeFileSync(join(root, 'model-source.js'), 'export const source = true;\n');
    const stats = {};
    for (const [flag, outDir] of [['true', enabledDir], ['false', disabledDir]]) {
      process.env.OOXML_MODULE_STATS = `${outDir}.json`;
      await build({ configFile: false, logLevel: 'silent', define: { __OOXML_MODEL_SOURCES__: flag },
        plugins: [moduleStatsPlugin(root)],
        build: { outDir, lib: { entry, formats: ['es'], fileName: () => 'entry.mjs' } } });
      stats[flag] = JSON.parse(readFileSync(`${outDir}.json`, 'utf8'));
    }
    delete process.env.OOXML_MODULE_STATS;
    assert.throws(() => assertDispatchCost(stats.true, stats.false, 'entry.mjs', 3_000, 'fixture'),
      /dispatch cost exceeds/);
    assert.ok(assertDispatchCost(stats.true, stats.false, 'entry.mjs', 100_000, 'fixture') > 3_000);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('dispatch cost ignores naming and partitioning, and a shrink never offsets growth', () => {
  const stats = (chunks) => ({ chunks });
  const disabled = stats({
    'entry.mjs': { imports: ['a.js', 'node:path'], modules: { 'src/entry.ts': 'const a = 1; f(a);' } },
    'a.js': { imports: [], modules: { 'src/shared.ts': 'function f(x) { return g(x, "shared"); }', 'src/other.ts': 'const o = [1, 2];' } },
  });
  // The same modules, renamed by chunk-scoped deconfliction and split across
  // other chunks with other placeholder ids, cost nothing.
  const repartitioned = stats({
    'entry.mjs': { imports: ['b.js', 'c.js'], modules: { 'src/entry.ts': 'const a$1 = 1; f$2(a$1);' } },
    'b.js': { imports: [], modules: { 'src/shared.ts': 'function f$2(x) { return g$1(x, "shared"); }' } },
    'c.js': { imports: [], modules: { 'src/other.ts': 'const o = [1, 2];' } },
  });
  assert.equal(assertDispatchCost(repartitioned, disabled, 'entry.mjs', 0, 'fixture'), 0);
  // Growth in one module is charged even when the same module shrinks.
  const grownAndShrunk = stats({
    'entry.mjs': { imports: ['a.js'], modules: { 'src/entry.ts': 'const a = 1; f(a);' } },
    'a.js': { imports: [], modules: { 'src/shared.ts': 'function f(x) { return "dispatch-only-table"; }', 'src/other.ts': 'const o = [1, 2];' } },
  });
  assert.equal(assertDispatchCost(grownAndShrunk, disabled, 'entry.mjs', 100, 'fixture'), Buffer.byteLength('"dispatch-only-table"'));
  // A module only the enabled graph retains counts whole.
  const added = stats({
    ...disabled.chunks,
    'entry.mjs': { imports: ['a.js'], modules: { 'src/entry.ts': 'const a = 1; f(a);', 'src/dispatch.ts': 'h("x");' } },
  });
  assert.ok(assertDispatchCost(added, disabled, 'entry.mjs', 100, 'fixture') > 0);
});

test('a source-only table hidden behind same-module shrinkage under repartitioning fails the budget', async () => {
  const root = mkdtempSync(join(tempRoot, 'ooxml-source-hide-'));
  try {
    const words = (prefix) => Array.from({ length: 80 }, (_, index) => `'${prefix}-${index}-${(index * 7919).toString(36)}'`).join(', ');
    writeFileSync(join(root, 'entry.js'), `import { shared } from './shared.js';
export async function load(options = {}) {
  if (__OOXML_MODEL_SOURCES__ && options.legacyDoc !== undefined) return import('./model-source.js');
  return shared(options.key);
}
`);
    // The enabled build retains a source-only table; the disabled build retains
    // a fallback of about the same size in the same module instead.
    writeFileSync(join(root, 'shared.js'), `const TABLE = __OOXML_MODEL_SOURCES__ ? [${words('source')}] : null;
const FALLBACK = __OOXML_MODEL_SOURCES__ ? null : [${words('fallbk')}];
export function shared(key) { return (TABLE ?? FALLBACK)[key]; }
`);
    // Importing the shared module from the source chunk moves it into its own
    // chunk in the enabled build only (repartitioning).
    writeFileSync(join(root, 'model-source.js'), "import { shared } from './shared.js';\nexport const source = shared(0);\n");
    const stats = {};
    for (const flag of ['true', 'false']) {
      const outDir = join(root, flag);
      process.env.OOXML_MODULE_STATS = `${outDir}.json`;
      await build({ configFile: false, logLevel: 'silent', define: { __OOXML_MODEL_SOURCES__: flag },
        plugins: [moduleStatsPlugin(root)],
        build: { outDir, minify: false, lib: { entry: join(root, 'entry.js'), formats: ['es'], fileName: () => 'entry.mjs' } } });
      stats[flag] = JSON.parse(readFileSync(`${outDir}.json`, 'utf8'));
    }
    delete process.env.OOXML_MODULE_STATS;
    assert.throws(() => assertDispatchCost(stats.true, stats.false, 'entry.mjs', 1_000, 'fixture'), /dispatch cost exceeds 1000-byte budget/);
  } finally {
    delete process.env.OOXML_MODULE_STATS;
    rmSync(root, { recursive: true, force: true });
  }
});
