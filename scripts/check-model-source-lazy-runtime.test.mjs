import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { runSourceRequestProbe } from './check-model-source-lazy-runtime.mjs';
import { checkSourceDispatchImports } from './check-model-source-lazy-graph.mjs';

const tempRoot = process.platform === 'darwin' ? '/private/tmp' : tmpdir();

test('synchronous getBuiltinModule/createRequire cannot bypass the runtime proof', () => {
  const temp = mkdtempSync(join(tempRoot, 'ooxml-source-cjs-probe-'));
  try {
    writeFileSync(join(temp, 'model-source-probe.cjs'), 'module.exports = 1;\n');
    const script = join(temp, 'probe.mjs');
    writeFileSync(script,
      "process.getBuiltinModule('module').createRequire(import.meta.url)('./model-source-probe.cjs');\n");
    assert.throws(() => runSourceRequestProbe({ script }),
      /Default OOXML load requested optional source code:[\s\S]*model-source-probe\.cjs/);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});

test('an inverted modelSources condition requests optional code on an ordinary load', () => {
  const temp = mkdtempSync(join(tempRoot, 'ooxml-inverted-source-probe-'));
  try {
    assert.doesNotThrow(() => checkSourceDispatchImports([{
      path: 'packages/xlsx/src/workbook.ts',
      text: "class XlsxWorkbook { static async load(options = {}) { if (options.modelSources === undefined) return import('./internal/workbook-model-source.js'); } }",
    }]));
    writeFileSync(join(temp, 'model-source-probe.mjs'), 'globalThis.sourceEvaluated = true;\n');
    writeFileSync(join(temp, 'entry.mjs'),
      "export async function load(options = {}) { if (options.modelSources === undefined) await import('./model-source-probe.mjs'); }\n");
    const script = join(temp, 'probe.mjs');
    writeFileSync(script, "import { load } from './entry.mjs'; await load();\n");
    assert.throws(() => runSourceRequestProbe({ script }),
      /Default OOXML load requested optional source code:[\s\S]*model-source-probe\.mjs/);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});

test('a renamed source module still fails when its runtime marker is loaded', () => {
  const temp = mkdtempSync(join(tempRoot, 'ooxml-renamed-source-probe-'));
  try {
    writeFileSync(join(temp, 'optional.mjs'),
      "export const marker = 'ooxml-model-source-module/v1';\n");
    const script = join(temp, 'probe.mjs');
    writeFileSync(script, "await import('./optional.mjs');\n");
    assert.throws(() => runSourceRequestProbe({ script }),
      /Default OOXML load requested optional source code:[\s\S]*optional\.mjs/);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});
