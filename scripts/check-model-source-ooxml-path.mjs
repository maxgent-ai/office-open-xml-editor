#!/usr/bin/env node
// Check ordinary OOXML awaits against a reviewed, versioned contract. A Git
// merge base rejects intentional upstream loader changes and becomes a no-op
// after merge. Update the baseline only after reviewing a loader's await order.
// Source-only branches are excluded. XLSX construction and parse must still
// share one host.run. This detects accidental edits, not hostile code.
// The DOCX baseline loads embedded fonts before Office fallbacks, then Google
// Fonts: embedded routes take precedence, and local Office faces avoid redundant
// Google substitutions. The other nine loader sequences remain unchanged.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';

function children(node) {
  return Object.values(node).flatMap((value) =>
    Array.isArray(value) ? value.filter((child) => child?.type)
      : value?.type ? [value] : []);
}

function findFunction(ast, name) {
  let found;
  function walk(node) {
    if (found) return;
    if ((node.type === 'FunctionDeclaration' && node.id?.name === name)
      || ((node.type === 'ClassMethod' || node.type === 'ClassPrivateMethod') && node.key?.name === name)) {
      found = node;
      return;
    }
    children(node).forEach(walk);
  }
  walk(ast.program);
  if (!found) throw new Error(`Missing function ${name}`);
  return found;
}

function callName(node) {
  if (!node) return '';
  if (node.type === 'CallExpression' || node.type === 'OptionalCallExpression') return callName(node.callee);
  if (node.type === 'MemberExpression' || node.type === 'OptionalMemberExpression') {
    return `${callName(node.object)}.${callName(node.property)}`;
  }
  if (node.type === 'Identifier') return node.name;
  if (node.type === 'Import') return 'import';
  if (node.type === 'NewExpression') return `new ${callName(node.callee)}`;
  return node.type;
}

function ooxmlAwaits(node, code) {
  const awaits = [];
  function walk(current) {
    if (current.type === 'BlockStatement') {
      for (const statement of current.body) {
        if (statement.type === 'IfStatement') {
          const test = code.slice(statement.test.start, statement.test.end).replace(/\s+/g, '');
          if (test === 'options.modelSources===undefined') {
            walk(statement.consequent);
            return;
          }
        }
        walk(statement);
      }
      return;
    }
    if (current.type === 'IfStatement') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      // The internal comparison-build flag can only narrow the existing
      // selected-source branch; the ordinary OOXML await path is unchanged.
      if (/^(?:__OOXML_MODEL_SOURCES__&&)?(?:opts|options)\.modelSources!==undefined$/.test(test)) {
        if (current.alternate) walk(current.alternate);
        return;
      }
      if (/^options\.modelSources===undefined$/.test(test)) {
        walk(current.consequent);
        return;
      }
      if (/^!sourceLoad$/.test(test)) {
        walk(current.consequent);
        return;
      }
    }
    if (current.type === 'ConditionalExpression') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      if (test === 'options.modelSources===undefined') {
        walk(current.consequent);
        return;
      }
      if (test === 'sourceLoad' || test === 'this._sourceLoad' || test.startsWith('sourceLoad&&')) {
        walk(current.alternate);
        return;
      }
    }
    if (current.type === 'AwaitExpression') awaits.push(callName(current.argument));
    children(current).forEach(walk);
  }
  walk(node.body);
  return awaits;
}

export function auditAwaitCase(file, name, baseline, current) {
  const candidate = ooxmlAwaits(findFunction(parse(current, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), current);
  if (JSON.stringify(candidate) !== JSON.stringify(baseline)) {
    throw new Error(`${file} ${name}: OOXML awaits changed\nbaseline ${JSON.stringify(baseline)}\nhead ${JSON.stringify(candidate)}`);
  }
  return candidate.length;
}

export function hasCombinedXlsxHostRun(code) {
  const ast = parse(code, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
  let found = false;
  function walk(node) {
    if (node.type === 'CallExpression' && callName(node.callee) === 'host.run') {
      const body = code.slice(node.start, node.end);
      if (/new XlsxArchive\(/.test(body) && /archive\.parse\(\)/.test(body)) found = true;
    }
    children(node).forEach(walk);
  }
  walk(ast.program);
  return found;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const cases = JSON.parse(readFileSync(new URL('./model-source-ooxml-awaits.json', import.meta.url), 'utf8'));
  for (const { file, name, awaits } of cases) {
    const count = auditAwaitCase(file, name, awaits, readFileSync(file, 'utf8'));
    console.log(`${file} ${name}: ${count} OOXML awaits, unchanged`);
  }
  const xlsxRender = readFileSync('packages/xlsx/src/render-worker.ts', 'utf8');
  if (!hasCombinedXlsxHostRun(xlsxRender)) {
    throw new Error('XLSX render worker splits OOXML construction and parse across host.run calls');
  }
  console.log('XLSX render worker keeps construction and parse in one host.run.');
}
