#!/usr/bin/env node
// Check ordinary OOXML awaits against a reviewed, versioned contract. A Git
// merge base rejects intentional upstream loader changes and becomes a no-op
// after merge. Update the baseline only after reviewing a loader's await order.
// Source-only branches are excluded. XLSX construction and parse must still
// share one host.run. This detects accidental edits, not hostile code.
// The DOCX baseline loads embedded fonts before Office fallbacks, then Google
// Fonts: embedded routes take precedence, and local Office faces avoid redundant
// Google substitutions. The PPTX font-owner transition is normalized only
// under the positive Google Fonts opt-in gate below.
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

function requiresGoogleFonts(node) {
  if (node.type === 'LogicalExpression' && node.operator === '&&') {
    return requiresGoogleFonts(node.left) || requiresGoogleFonts(node.right);
  }
  return node.type === 'MemberExpression' && !node.computed
    && node.object.type === 'Identifier' && node.object.name === 'opts'
    && node.property.type === 'Identifier' && node.property.name === 'useGoogleFonts';
}

function ooxmlAwaits(node, code, pptxFontOwnerRefactor = false) {
  const awaits = [];
  function walk(current, googleFontsRequired = false) {
    if (current.type === 'BlockStatement') {
      for (const statement of current.body) {
        if (statement.type === 'IfStatement') {
          const test = code.slice(statement.test.start, statement.test.end).replace(/\s+/g, '');
          if (test === 'options.modelSources===undefined') {
            walk(statement.consequent, googleFontsRequired);
            return;
          }
        }
        walk(statement, googleFontsRequired);
      }
      return;
    }
    if (current.type === 'IfStatement') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      // The internal comparison-build flag can only narrow the existing
      // selected-source branch; the ordinary OOXML await path is unchanged.
      if (/^(?:__OOXML_MODEL_SOURCES__&&)?(?:opts|options)\.modelSources!==undefined$/.test(test)) {
        if (current.alternate) walk(current.alternate, googleFontsRequired);
        return;
      }
      if (/^options\.modelSources===undefined$/.test(test)) {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      if (/^!sourceLoad$/.test(test)) {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      walk(current.test, googleFontsRequired);
      walk(current.consequent, googleFontsRequired || requiresGoogleFonts(current.test));
      if (current.alternate) walk(current.alternate, googleFontsRequired);
      return;
    }
    if (current.type === 'ConditionalExpression') {
      const test = code.slice(current.test.start, current.test.end).replace(/\s+/g, '');
      if (test === 'options.modelSources===undefined') {
        walk(current.consequent, googleFontsRequired);
        return;
      }
      if (test === 'sourceLoad' || test === 'this._sourceLoad' || test.startsWith('sourceLoad&&')) {
        walk(current.alternate, googleFontsRequired);
        return;
      }
    }
    if (current.type === 'AwaitExpression') {
      let name = callName(current.argument);
      // The PPTX font barrier moved from a per-call array owner to the shared
      // presentation lease. Normalize only this known callee transition under
      // its positive opt-in AND gate. Every await is still visited, so additions
      // or reordering inside the gate fail; unguarded/OR/alternate calls and
      // every other package/helper retain their original identity checks.
      if (pptxFontOwnerRefactor && googleFontsRequired && name === 'pres._ensureGoogleFonts') {
        name = 'preloadGoogleFonts';
      }
      awaits.push(name);
    }
    children(current).forEach(child => walk(child, googleFontsRequired));
  }
  walk(node.body);
  return awaits;
}

export function auditAwaitCase(file, name, baseline, current) {
  const pptxFontOwnerRefactor = file === 'packages/pptx/src/presentation.ts' && name === 'load';
  const candidate = ooxmlAwaits(findFunction(parse(current, { sourceType: 'module', plugins: ['typescript', 'jsx'] }), name), current, pptxFontOwnerRefactor);
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
