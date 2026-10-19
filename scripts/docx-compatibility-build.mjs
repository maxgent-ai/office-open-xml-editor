import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { checkDocxCompatibilityEvidence } from './check-docx-compatibility-evidence.mjs';

const require = createRequire(import.meta.url);
const ts = require('typescript-compiler-api');
const root = resolve(import.meta.dirname, '..');
// Filesystem reads keep native paths; all module-ID comparisons use Vite's
// forward-slash spelling on Windows as well as POSIX hosts.
const modulePath = (file) => file.replaceAll('\\', '/');

function literalMembers(object, names, file) {
  if (!object || !ts.isObjectLiteralExpression(object)) {
    throw new Error(`${file}: compatibility metadata must be an object literal`);
  }
  const members = new Map();
  for (const member of object.properties) {
    if (!ts.isPropertyAssignment(member)
      || !(ts.isIdentifier(member.name) || ts.isStringLiteral(member.name))
      || !names.includes(member.name.text) || members.has(member.name.text)) {
      throw new Error(`${file}: unsupported compatibility member`);
    }
    members.set(member.name.text, member.initializer);
  }
  if (members.size !== names.length) throw new Error(`${file}: incomplete compatibility metadata`);
  return members;
}

/**
 * Source evidence remains beside its implementation and is checked before any
 * build. Published renderers consume only stable IDs. Replace checked factory
 * calls, rather than deleting arbitrary properties or weakening the source
 * factory: Vitest/dev retain full validation, and production retains frozen IDs.
 * Literal-only metadata is essential here: stripping must never discard effects.
 */
export function stripDocxCompatibilityMetadata(code, file) {
  const source = ts.createSourceFile(file, code, ts.ScriptTarget.Latest, true);
  if (source.parseDiagnostics.length) throw new Error(`${file}: invalid compatibility source`);
  const edits = [];
  const descriptions = [];
  function visit(node) {
    if (ts.isCallExpression(node) && ts.isIdentifier(node.expression)
      && node.expression.text === 'defineCompatibilityRule') {
      if (node.arguments.length !== 1) throw new Error(`${file}: unsupported compatibility call`);
      const members = literalMembers(node.arguments[0], ['id', 'evidence', 'description'], file);
      for (const name of ['id', 'description']) {
        if (!ts.isStringLiteralLike(members.get(name))) throw new Error(`${file}: nonliteral ${name}`);
      }
      const id = members.get('id');
      if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(id.text)) throw new Error(`${file}: invalid rule ID`);
      const evidence = members.get('evidence');
      const kind = evidence && ts.isObjectLiteralExpression(evidence)
        ? evidence.properties.find((member) => ts.isPropertyAssignment(member)
          && member.name.text === 'kind')?.initializer : undefined;
      const fields = kind && ts.isStringLiteralLike(kind)
        && ['microsoft-note', 'regression-test', 'library-policy'].includes(kind.text)
        ? ['kind', 'reference']
        : kind && ts.isStringLiteralLike(kind) && kind.text === 'office-observation'
          ? ['kind', 'syntheticFixtureId', 'application', 'version', 'platform'] : [];
      if (!fields.length) throw new Error(`${file}: unsupported evidence kind`);
      for (const value of literalMembers(evidence, fields, file).values()) {
        if (!ts.isStringLiteralLike(value)) throw new Error(`${file}: nonliteral evidence`);
      }
      descriptions.push(members.get('description').text);
      edits.push({ start: node.getStart(source), end: node.end,
        text: `/* @__PURE__ */ Object.freeze({ id: ${id.getText(source)} })` });
      return;
    }
    ts.forEachChild(node, visit);
  }
  visit(source);
  for (const edit of edits.reverse()) {
    code = code.slice(0, edit.start) + edit.text + code.slice(edit.end);
  }
  return { code, descriptions };
}

/** Decode literal escapes so minification cannot hide leaked descriptions. */
export function assertNoDocxCompatibilityDescriptions(code, descriptions, file) {
  const source = ts.createSourceFile(file, code, ts.ScriptTarget.Latest, false, ts.ScriptKind.JS);
  function visit(node) {
    if (ts.isStringLiteralLike(node) && descriptions.has(node.text)) {
      throw new Error(`${file}: bundled DOCX compatibility description`);
    }
    ts.forEachChild(node, visit);
  }
  visit(source);
}

export function docxCompatibilityMetadata() {
  let files;
  let descriptions;
  function prepare() {
    // Also rejects factory aliases/re-exports and declarations outside the
    // source authority, so a new spelling or module cannot silently escape.
    const sourceFiles = new Set(checkDocxCompatibilityEvidence(root).map((rule) => resolve(root, rule.file)));
    files = new Set([...sourceFiles].map(modulePath));
    descriptions = new Set();
    for (const file of sourceFiles) {
      for (const description of stripDocxCompatibilityMetadata(readFileSync(file, 'utf8'), file).descriptions) {
        descriptions.add(description);
      }
    }
  }
  return {
    name: 'docx-compatibility-metadata',
    enforce: 'pre',
    apply: 'build',
    buildStart() {
      files = undefined;
      descriptions = undefined;
    },
    transform(code, id) {
      const file = modulePath(id.split('?')[0]);
      // Defer source validation until a rule owner is visited. It necessarily
      // contains the factory name even if a caller aliases it; prepare() then
      // checks the entire authority before transforming any rule module.
      if (!files && file.startsWith(modulePath(resolve(root, 'packages/docx/src')) + '/')
        && code.includes('defineCompatibilityRule')) prepare();
      if (!files?.has(file)) return null;
      return { code: stripDocxCompatibilityMetadata(code, id).code, map: null };
    },
    generateBundle: {
      // Vite emits worker assets in its normal generateBundle hook, after pre
      // plugins. Validate last, including worker-only graphs whose main bundle
      // never transformed a rule module, and decode binary asset sources too.
      order: 'post',
      handler(_options, bundle) {
        if (!descriptions) prepare();
        for (const output of Object.values(bundle)) {
          if (output.type === 'chunk') {
            assertNoDocxCompatibilityDescriptions(output.code, descriptions, output.fileName);
          } else if (/\.[cm]?js$/.test(output.fileName)) {
            const code = typeof output.source === 'string'
              ? output.source : new TextDecoder().decode(output.source);
            assertNoDocxCompatibilityDescriptions(code, descriptions, output.fileName);
          }
        }
      },
    },
  };
}
