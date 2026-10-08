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
  const sourceOnly = 'async function load() { if (opts.modelSources !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, sourceOnly), 2);
  const gatedSourceOnly = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.modelSources !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, gatedSourceOnly), 2);
  const inverted = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.modelSources === undefined) await helper(); await open(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, inverted), /OOXML awaits changed/);
});

test('XLSX construction and parse must share one host.run', () => {
  assert.equal(hasCombinedXlsxHostRun('host.run(() => { const archive = new XlsxArchive(bytes); return archive.parse(); });'), true);
  assert.equal(hasCombinedXlsxHostRun('host.run(() => new XlsxArchive(bytes)); host.run(() => archive.parse());'), false);
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

test('PPTX retains the same opt-in font barrier when its owner becomes a lease', () => {
  const file = 'packages/pptx/src/presentation.ts';
  const previous = ['parse', 'preloadGoogleFonts'];
  const changed = 'async function load() { await parse(); if (mode === "main" && opts.useGoogleFonts && ready) await pres._ensureGoogleFonts(names); }';
  assert.equal(auditAwaitCase(file, 'load', previous, changed), 2);
  const unguarded = 'async function load() { await parse(); await pres._ensureGoogleFonts(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, unguarded), /OOXML awaits changed/);
  const extra = 'async function load() { await parse(); if (opts.useGoogleFonts) { await pres._ensureGoogleFonts(names); await helper(); } }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, extra), /OOXML awaits changed/);
  const optional = 'async function load() { await parse(); if (mode === "main" || opts.useGoogleFonts) await pres._ensureGoogleFonts(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, optional), /OOXML awaits changed/);
});
