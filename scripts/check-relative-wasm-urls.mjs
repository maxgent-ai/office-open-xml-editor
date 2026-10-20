#!/usr/bin/env node
// Guard that every built bundle locates its parser WASM relative to the
// module (or CommonJS file) that loads it, never relative to the hosting
// page's origin root. A root-absolute "/x.wasm" reference works only when the
// package is served at the site root: below a subpath or from a CDN the fetch
// returns 404. Vite 8.3 emits that form for library builds without
// `base: './'` (#1893 review).
//
// Usage: node scripts/check-relative-wasm-urls.mjs <dist-dir> [...]
// Each directory must reference at least one .wasm asset that exists beside
// the bundles. Every .wasm string literal in its .js/.mjs/.cjs files must be
//   - a bare file name resolved against the module: within the same
//     expression as import.meta.url, document.currentScript, __filename or
//     __dirname (the forms Rollup/Rolldown emit for ES and CommonJS output); or
//   - a "/name.wasm" suffix appended to __dirname (CommonJS file paths).
// Source-module ids such as "src/wasm/x.wasm?url" in comments are ignored.

import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { basename, extname, join } from 'node:path';

const JS_EXT = new Set(['.js', '.mjs', '.cjs']);
const LITERAL = /(["'`])([^"'`\s]*\.wasm)\1/g;
const MODULE_BASES = ['import.meta.url', 'document.currentScript', '__filename', '__dirname'];
const CONTEXT_CHARS = 240;

/** @param {string} dir @param {string[]} out */
function collect(dir, out) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) collect(path, out);
    else out.push(path);
  }
}

/**
 * Classify one literal. Exported shape is kept simple so the unit test can
 * exercise the rule on synthetic bundle text.
 * @param {string} text @param {number} index offset of the opening quote @param {string} value
 * @returns {string | null} a problem description, or null when relative
 */
export function wasmReferenceProblem(text, index, value) {
  if (value.includes('?')) return null;
  if (/^[a-z][a-z0-9+.-]*:/i.test(value)) return `absolute URL ${value}`;
  if (value.startsWith('/')) {
    return /__dirname\s*\+\s*$/.test(text.slice(Math.max(0, index - 40), index))
      ? null : `root-absolute URL ${value}`;
  }
  const context = text.slice(Math.max(0, index - CONTEXT_CHARS), index + value.length + CONTEXT_CHARS);
  return MODULE_BASES.some((base) => context.includes(base))
    ? null : `${value} is not resolved against the loading module`;
}

/** @param {string} dir */
export function checkDist(dir) {
  const files = [];
  collect(dir, files);
  const wasmFiles = new Set(files.filter((file) => extname(file) === '.wasm').map((file) => basename(file)));
  const problems = [];
  let references = 0;
  for (const file of files) {
    if (!JS_EXT.has(extname(file))) continue;
    const text = readFileSync(file, 'utf8');
    for (const match of text.matchAll(LITERAL)) {
      const value = match[2];
      if (value.includes('?')) continue;
      references++;
      const problem = wasmReferenceProblem(text, match.index, value);
      if (problem) problems.push(`${file}: ${problem}`);
      const name = basename(value);
      if (!wasmFiles.has(name)) problems.push(`${file}: references ${name}, which is not in ${dir}`);
    }
  }
  if (references === 0) problems.push(`${dir}: no .wasm reference found`);
  return { problems, references };
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const dirs = process.argv.slice(2);
  if (dirs.length === 0) {
    console.error('usage: check-relative-wasm-urls.mjs <dist-dir> [...]');
    process.exit(2);
  }
  let failed = false;
  for (const dir of dirs) {
    if (!existsSync(dir)) {
      console.error(`check-relative-wasm-urls: directory not found: ${dir}`);
      process.exit(2);
    }
    const { problems, references } = checkDist(dir);
    for (const problem of problems) console.error(`  ${problem}`);
    console.log(`check-relative-wasm-urls: ${dir}: ${references} reference(s), ${problems.length} problem(s)`);
    failed ||= problems.length > 0;
  }
  process.exit(failed ? 1 : 0);
}
