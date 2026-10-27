import assert from 'node:assert/strict';
import test from 'node:test';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { auditAwaitCase, hasCombinedXlsxHostRun } from './check-model-source-ooxml-path.mjs';

test('an OOXML await added outside a selected-source branch fails the AST audit', () => {
  const previous = ['open', 'parse'];
  const changed = 'async function load() { await open(); await helper(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, changed), /OOXML awaits changed/);
  const reordered = 'async function load() { await parse(); await open(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, reordered), /OOXML awaits changed/);
  const removed = 'async function load() { await open(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, removed), /OOXML awaits changed/);
  const sourceOnly = 'async function load() { if (opts.legacyDoc !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, sourceOnly), 2);
  const gatedSourceOnly = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.legacyDoc !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, gatedSourceOnly), 2);
  const inverted = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.legacyDoc === undefined) await helper(); await open(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, inverted), /OOXML awaits changed/);
});

test('XLSX construction and parse must share one host.run', () => {
  assert.equal(hasCombinedXlsxHostRun('host.run(() => { const archive = new XlsxArchive(bytes); return archive.parse(); });'), true);
  assert.equal(hasCombinedXlsxHostRun('host.run(() => new XlsxArchive(bytes)); host.run(() => archive.parse());'), false);
});

test('the renamed per-format source branches preserve the ordinary await contract', () => {
  for (const key of ['modelSources', 'legacyDoc', 'legacyXls', 'legacyPpt']) {
    const baseline = ['open', 'parse'];
    const selected = `async function load() { if (opts.${key} !== undefined) await helper(); await open(); await parse(); }`;
    const absent = `async function load() { if (options.${key} === undefined) { await open(); await parse(); return; } await helper(); }`;
    assert.equal(auditAwaitCase('sample.ts', 'load', baseline, selected), 2);
    assert.equal(auditAwaitCase('sample.ts', 'load', baseline, absent), 2);
    const unguarded = `async function load() { if (options.${key} === undefined) { await open(); await helper(); await parse(); return; } }`;
    assert.throws(() => auditAwaitCase('sample.ts', 'load', baseline, unguarded), /OOXML awaits changed/);
  }
});

// A reviewed contract must work in source archives and shallow CI checkouts.
test('the CLI checks the reviewed await baseline without Git history', () => {
  const result = spawnSync(process.execPath, ['scripts/check-model-source-ooxml-path.mjs'], {
    cwd: fileURLToPath(new URL('../', import.meta.url)),
    env: { ...process.env, PATH: '' },
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, result.stdout + result.stderr);
  assert.match(result.stdout, /XLSX render worker keeps construction and parse/);
});

test('PPTX keeps the reviewed font-source await behind its positive opt-in gate', () => {
  const file = 'packages/pptx/src/presentation.ts';
  const previous = ['parse', 'pres._ensureFontSources'];
  const changed = 'async function load() { await parse(); if (mode === "main" && opts.fonts && ready) await pres._ensureFontSources(names); }';
  assert.equal(auditAwaitCase(file, 'load', previous, changed), 2);
  const unguarded = 'async function load() { await parse(); await pres._ensureFontSources(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, unguarded), /positive opts.fonts gate/);
  const extra = 'async function load() { await parse(); if (opts.fonts) { await pres._ensureFontSources(names); await helper(); } }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, extra), /OOXML awaits changed/);
  const optional = 'async function load() { await parse(); if (mode === "main" || opts.fonts) await pres._ensureFontSources(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, optional), /positive opts.fonts gate/);
  const alternate = 'async function load() { await parse(); if (opts.fonts) {} else await pres._ensureFontSources(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, alternate), /positive opts.fonts gate/);
  const moved = 'async function load() { if (opts.fonts) await pres._ensureFontSources(names); await parse(); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, moved), /OOXML awaits changed/);
  const twice = 'async function load() { await parse(); if (opts.fonts) { await pres._ensureFontSources(names); await pres._ensureFontSources(names); } }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, twice), /OOXML awaits changed/);
});

test('DOCX requires the reviewed native route barrier after math preparation', () => {
  const file = 'packages/docx/src/document.ts';
  const baseline = ['parse', 'prepareMathRuns', 'nativeRoutesPending', 'publish'];
  const current = 'async function load() { await parse(); await prepareMathRuns(); await nativeRoutesPending; await publish(); }';
  assert.equal(auditAwaitCase(file, 'load', baseline, current), 4);
  for (const changed of [
    'async function load() { await nativeRoutesPending; await parse(); await prepareMathRuns(); await publish(); }',
    'async function load() { await parse(); await prepareMathRuns(); await publish(); }',
    'async function load() { await parse(); await prepareMathRuns(); await nativeRoutesPending; await nativeRoutesPending; await publish(); }',
  ]) {
    assert.throws(() => auditAwaitCase(file, 'load', baseline, changed), /OOXML awaits changed/);
  }
});
