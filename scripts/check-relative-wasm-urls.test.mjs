import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { checkDist } from './check-relative-wasm-urls.mjs';

function dist(files) {
  const dir = mkdtempSync(join(tmpdir(), 'ooxml-wasm-url-check-'));
  writeFileSync(join(dir, 'x_parser_bg.wasm'), '');
  for (const [name, text] of Object.entries(files)) writeFileSync(join(dir, name), text);
  return dir;
}

function problems(files) {
  const dir = dist(files);
  try { return checkDist(dir).problems; } finally { rmSync(dir, { recursive: true, force: true }); }
}

test('accepts the module-relative forms Rolldown emits for ES and CommonJS output', () => {
  assert.deepEqual(problems({
    'a.mjs': '//#region src/wasm/x_parser_bg.wasm?url\nvar u = new URL("x_parser_bg.wasm", import.meta.url).href;',
    'b.cjs': 'var u = new URL(`x_parser_bg.wasm`,require("node:url").pathToFileURL(__filename)).href;',
    'c.cjs': 'var u=typeof document>`u`?require("url").pathToFileURL(__dirname+`/x_parser_bg.wasm`).href'
      + ':new URL(`x_parser_bg.wasm`,document.currentScript&&document.currentScript.src).href;',
  }), []);
});

test('rejects the root-absolute reference a library build emits without base ./', () => {
  const found = problems({ 'a.js': 'var u = "/x_parser_bg.wasm";' });
  assert.equal(found.length, 1);
  assert.match(found[0], /root-absolute URL \/x_parser_bg\.wasm/);
});

test('rejects page-relative, absolute and missing references', () => {
  assert.match(problems({ 'a.js': 'fetch("x_parser_bg.wasm");' })[0], /not resolved against the loading module/);
  assert.match(problems({ 'a.js': 'fetch("https://cdn.test/x_parser_bg.wasm", import.meta.url);' })[0], /absolute URL/);
  assert.match(problems({ 'a.mjs': 'new URL("y_parser_bg.wasm", import.meta.url);' })[0], /not in/);
  assert.match(problems({ 'a.mjs': 'export {};' })[0], /no \.wasm reference/);
});
