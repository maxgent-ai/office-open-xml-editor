import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createTypeScriptResolver, findViolations, isGuardedPath } from './check-core-legacy-boundary.mjs';
import { legacyBundleBoundary } from '../vite.config.ts';
import { build } from 'vite';

const rules = (text) => findViolations([{ path: 'packages/docx/src/x.ts', text }]).map((v) => v.rule);

test('flags legacy imports, options, branches and names in the OOXML packages', () => {
  assert.deepEqual(rules("import { x } from '@silurus/ooxml-legacy-converter/internal/direct-doc-engine';"),
    ['legacy-package', 'direct-format-name']);
  assert.deepEqual(rules('await DocxDocument.load(bytes, { legacyConversion: {} });'), ['legacy-office-api']);
  assert.deepEqual(rules("if (kind === 'legacy-xls') return;"), ['legacy-format-name']);
  assert.deepEqual(rules('const nativeDoc = await openNative();'), ['native-legacy-source']);
  assert.deepEqual(rules('archive.revision_markup_in_print?.()'), ['legacy-revision-view']);
  assert.deepEqual(rules('interface T { measureLegacyXlsNormalFont?: unknown; }'), ['legacy-xls-measurement']);
  assert.deepEqual(rules('interface T { langDefault?: string; }'), ['docx-lang-default']);
});

test('allows the pre-existing generic CFB rejection and unrelated "legacy" wording', () => {
  assert.deepEqual(rules("throw new OoxmlError('legacy-binary-format', 'a legacy binary .doc file');"), []);
  assert.deepEqual(rules('// legacy VML shapes and the legacy chart family fallback'), []);
  assert.deepEqual(rules('export function selectModelSource(sources, target, bytes) {}'), []);
});

test('decodes escaped static, dynamic, require and re-export module specifiers', () => {
  for (const code of [
    "import x from '@silurus/ooxml-\\u006cegacy-converter';",
    "export * from '@silurus/ooxml-\\u006cegacy-converter';",
    "await import('@silurus/ooxml-\\u006cegacy-converter');",
    "require('@silurus/ooxml-\\u006cegacy-converter');",
    "import x from '../../legacy-\\u0063onverter/src/index.js';",
  ]) {
    assert.ok(rules(code).includes('legacy-package'), code);
  }
});

test('rejects computed imports and requires, import attributes and type-level imports', () => {
  for (const code of [
    "import('@silurus/ooxml-' + 'legacy' + '-converter');",
    "import(`@silurus/ooxml-${'legacy'}-converter`);",
    "const p = '@silurus/ooxml-\\u006cegacy-converter'; require(p);",
  ]) {
    assert.ok(rules(code).includes('computed-module'), code);
  }
  assert.deepEqual(rules("import('@silurus/ooxml-\\u006cegacy-converter/package.json', { with: { type: 'json' } });"), ['legacy-package']);
  assert.deepEqual(rules("type X = import('@silurus/ooxml-\\u006cegacy-converter').X;"), ['legacy-package']);
  assert.ok(rules('require(variable);').includes('computed-module'));
  assert.ok(rules('require(variable);').includes('loader-identifier'));
  assert.deepEqual(rules('import(variable);'), ['computed-module']);
  assert.deepEqual(
    findViolations([{ path: 'packages/docx/src/new-probe.test.ts', text: 'import(variable);' }]).map((v) => v.rule),
    ['computed-module'],
  );
});

test('rejects indirect Node require forms even with escaped module names', () => {
  for (const code of [
    "import { createRequire } from 'node:module'; const load = createRequire(import.meta.url); load(name);",
    "import { createRequire as cr } from 'node:module'; cr(import.meta.url)(name);",
    'require.call(null, name);',
    'require.apply(null, [name]);',
    'require.bind(null)(name);',
    'require?.(name);',
    'module.require(name);',
    "module['require'](name);",
    "module?.['require'](name);",
    'module?.require?.(name);',
  ]) {
    assert.ok(rules(code).some((rule) => rule === 'node-module-import' || rule === 'loader-identifier'), code);
  }
});

test('rejects every Node module import form and loader identifier reference', () => {
  for (const code of [
    "import { 'createRequire' as load } from 'node:module'; load(import.meta.url)(name);",
    "import Module from 'node:module'; const { 'createRequire': load } = Module; load(import.meta.url)(name);",
    "import Module from 'node:module'; Module['create' + 'Require'](import.meta.url)(name);",
    "import * as loader from 'module';",
    "import('node:module');",
    "const { module } = value;",
    "module['require'](name);",
    'require?.(name);',
  ]) {
    assert.ok(rules(code).some((rule) => rule === 'node-module-import' || rule === 'loader-identifier'), code);
  }
});

test('rejects alternate runtime loaders and code generation capabilities', () => {
  for (const code of [
    "process.getBuiltinModule('module').createRequire(import.meta.url)(name);",
    "globalThis['process'].getBuiltinModule('module').createRequire(import.meta.url)(name);",
    "const load = process['getBuiltinModule']('module');",
    "process.binding('natives');",
    "globalThis.process.binding('natives');",
    "eval('process');",
    "(0, eval)('process');",
    "globalThis.eval('process');",
    "new Function('return process')();",
    "Function('return process')();",
    "globalThis.Function('return process')();",
    "const make = Function; make('return process')();",
    "const make = globalThis['Function']; make('return process')();",
    "({}).constructor.constructor('return process')();",
    "importScripts(path);",
    "const load = importScripts; load(path);",
    "const load = self['importScripts']; load(path);",
    "globalThis[key];",
    "self[key];",
    "window[key];",
    "globalThis['eval']('process');",
    "globalThis['process']['binding']('natives');",
  ]) {
    assert.ok(rules(code).length > 0, code);
  }
  for (const code of [
    "importScripts('./worker.js');",
    "globalThis['document'];",
    "self[`location`];",
    "window[0];",
  ]) {
    assert.deepEqual(rules(code), [], code);
  }
  assert.deepEqual(findViolations([{
    path: 'packages/core/src/worker/bridge.test.ts', text: 'expect.any(Function);',
  }]), []);
});

test('follows a relay outside guarded roots to a dynamic legacy import', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-relay-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-converter/src'), { recursive: true });
    mkdirSync(join(root, 'shared'), { recursive: true });
    writeFileSync(join(root, 'packages/core/tsconfig.json'), JSON.stringify({
      compilerOptions: { moduleResolution: 'bundler', module: 'esnext' },
    }));
    writeFileSync(join(root, 'packages/legacy-converter/src/index.ts'), 'export const reader = 1;');
    writeFileSync(join(root, 'shared/reader.ts'),
      "export const reader = () => import('../packages/legacy-converter/src/index.ts');");
    const violations = findViolations([{
      path: 'packages/core/src/probe.ts',
      text: "export { reader } from '../../../shared/reader.ts';",
    }], { resolveModule: createTypeScriptResolver(root) });
    assert.ok(violations.some((violation) => violation.rule === 'legacy-package'
      && violation.path === 'shared/reader.ts'));
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('TypeScript module resolution catches a path alias into the legacy package', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-converter/src'), { recursive: true });
    writeFileSync(join(root, 'packages/legacy-converter/src/index.ts'), 'export const reader = 1;');
    writeFileSync(join(root, 'packages/core/tsconfig.json'), JSON.stringify({
      compilerOptions: { baseUrl: '../..', paths: { '@reader': ['packages/legacy-converter/src/index.ts'] }, moduleResolution: 'bundler', module: 'esnext' },
    }));
    const violations = findViolations(
      [{ path: 'packages/core/src/probe.ts', text: "import { reader } from '@reader';" }],
      { resolveModule: createTypeScriptResolver(root) },
    );
    assert.deepEqual(violations.map((v) => v.rule), ['legacy-package']);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('built chunk graph rejects a resolved legacy module even after aliasing', () => {
  const hook = legacyBundleBoundary().generateBundle;
  assert.equal(typeof hook, 'function');
  const invoke = (moduleId, fileName = 'entry.js') => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    { [fileName]: { type: 'chunk', fileName, isEntry: true, imports: [], modules: { [moduleId]: {} } } },
  );
  assert.doesNotThrow(() => invoke('/repo/packages/core/src/index.ts'));
  assert.throws(
    () => invoke('/repo/packages/legacy-converter/src/index.ts'),
    /forbidden legacy module/,
  );
  assert.doesNotThrow(() => invoke('/repo/packages/legacy-converter/src/index.ts', 'legacy-ppt.mjs'));
  assert.throws(() => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    {
      'docx.mjs': { type: 'chunk', fileName: 'docx.mjs', isEntry: true, imports: ['shared.js'], modules: {} },
      'shared.js': { type: 'chunk', fileName: 'shared.js', isEntry: false, imports: [], modules: { '/repo/packages/legacy-converter/src/index.ts': {} } },
    },
  ), /forbidden legacy module/);
  assert.throws(() => hook.call(
    { error(message) { throw new Error(message); } },
    {},
    {
      'docx.mjs': { type: 'chunk', fileName: 'docx.mjs', isEntry: true, imports: [], dynamicImports: ['relay.js'], modules: {} },
      'relay.js': { type: 'chunk', fileName: 'relay.js', isEntry: false, imports: [], dynamicImports: [], modules: { '/repo/packages/legacy-converter/src/index.ts': {} } },
    },
  ), /forbidden legacy module/);
});

test('a real build rejects a dynamic chunk reached through a relay', async () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-boundary-bundle-'));
  try {
    mkdirSync(join(root, 'packages/core/src'), { recursive: true });
    mkdirSync(join(root, 'packages/legacy-converter/src'), { recursive: true });
    mkdirSync(join(root, 'shared'), { recursive: true });
    const entry = join(root, 'packages/core/src/probe.ts');
    writeFileSync(entry, "export { reader } from '../../../shared/reader.ts';");
    writeFileSync(join(root, 'shared/reader.ts'),
      "export const reader = () => import('../packages/legacy-converter/src/index.ts');");
    writeFileSync(join(root, 'packages/legacy-converter/src/index.ts'), 'export const value = 1;');
    await assert.rejects(() => build({
      configFile: false,
      logLevel: 'silent',
      plugins: [legacyBundleBoundary()],
      build: { write: false, lib: { entry, formats: ['es'] } },
    }), /forbidden legacy module/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('guards source, parser and manifest files but not generated or private output', () => {
  assert.equal(isGuardedPath('packages/core/src/index.ts'), true);
  assert.equal(isGuardedPath('packages/docx/parser/src/parser.rs'), true);
  assert.equal(isGuardedPath('packages/node/package.json'), true);
  assert.equal(isGuardedPath('packages/xlsx/src/wasm/xlsx_parser.js'), false);
  assert.equal(isGuardedPath('packages/pptx/public/demo/readme.md'), false);
  assert.equal(isGuardedPath('packages/legacy-converter/src/legacy-doc.ts'), false);
});
