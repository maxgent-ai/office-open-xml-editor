#!/usr/bin/env node
// Generate packages/ooxml-common/src/spreadsheet_style_presets.generated.rs
// from ECMA-376 Part 1 Annex G presetTableStyles.xml: the built-in Table and
// PivotTable styles (§18.8.40) as compact dictionary indices, decoded on lookup.
// Usage: node scripts/generate-spreadsheet-style-presets.mjs <presetTableStyles.xml>
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const source = process.argv[2];
if (!source) throw new Error('usage: generate-spreadsheet-style-presets.mjs <presetTableStyles.xml>');
const xml = readFileSync(source, 'utf8');

const attr = (tag, name) => {
  const match = new RegExp(`\\b${name}="([^"]*)"`).exec(tag);
  return match ? match[1] : null;
};
const color = (tag) => {
  if (!tag) return null;
  if (attr(tag, 'rgb') || attr(tag, 'indexed') || attr(tag, 'auto')) throw new Error(`unsupported color ${tag}`);
  const theme = attr(tag, 'theme');
  if (theme === null) throw new Error(`color without theme ${tag}`);
  const tint = attr(tag, 'tint');
  if (!Number.isInteger(Number(theme)) || Number(theme) < 0 || Number(theme) > 255) {
    throw new Error(`theme index exceeds u8: ${theme}`);
  }
  if (tint !== null && !Number.isFinite(Number(tint))) throw new Error(`invalid tint: ${tint}`);
  return [Number(theme), tint === null ? null : Number(tint)];
};
const known = /^(dxf|font|b|color|fill|patternFill|fgColor|bgColor|border|left|right|top|bottom|horizontal|vertical)$/;

/** One `<dxf>` as structured data, preserving absent versus present components. */
function dxfExpr(dxf) {
  for (const [, name] of dxf.matchAll(/<\/?(\w+)/g)) {
    if (!known.test(name)) throw new Error(`unsupported dxf element ${name} in ${dxf}`);
  }
  // PresetFont.bold means "<b/> present"; any other <b> form (val="0", ...)
  // would need the absent/explicit-off distinction the preset type lacks.
  if (/<b\s[^>]*>/.test(dxf)) throw new Error(`unsupported bold form in ${dxf}`);
  const font = /<font>([\s\S]*?)<\/font>/.exec(dxf);
  const fontExpr = font
    ? [/<b\/>/.test(font[1]), color(/<color [^>]*\/>/.exec(font[1])?.[0])]
    : null;
  const pattern = /<patternFill([^>]*)>([\s\S]*?)<\/patternFill>/.exec(dxf);
  const fillExpr = pattern
    ? [attr(pattern[1], 'patternType') ?? 'solid', color(/<fgColor [^>]*\/>/.exec(pattern[2])?.[0]), color(/<bgColor [^>]*\/>/.exec(pattern[2])?.[0])]
    : null;
  if (/<fill>/.test(dxf) && !pattern) throw new Error(`unsupported fill ${dxf}`);
  const border = /<border>([\s\S]*?)<\/border>/.exec(dxf);
  let borderExpr = null;
  if (border) {
    const edges = {};
    for (const [, name, attrs, body] of border[1].matchAll(/<(\w+)([^>]*?)(?:\/>|>([\s\S]*?)<\/\1>)/g)) {
      const style = attr(attrs, 'style');
      edges[name] = [style, color(body && /<color [^>]*\/>/.exec(body)?.[0])];
    }
    const edge = (name) => edges[name] ?? null;
    borderExpr = ['left', 'right', 'top', 'bottom', 'horizontal', 'vertical'].map(edge);
  }
  return [fontExpr, fillExpr, borderExpr];
}

const dxfIds = new Map();
// Annex G numbers each style's dxfIds from 0 for PivotTable styles and from 1
// for Table styles (every TableStyle* block runs 1..count over its <dxfs>);
// the base is checked per block so a mismatch fails generation.
function readStyles(family, base) {
  const out = [];
  for (const [, name, raw] of xml.matchAll(new RegExp(`<(${family}(?:Light|Medium|Dark)\\d+)>([\\s\\S]*?)<\\/\\1>`, 'g'))) {
    const body = raw.trim().replace(/>\s+</g, '><');
    const local = [...body.matchAll(/<dxf>[\s\S]*?<\/dxf>/g)].map(([dxf]) => {
      const expr = JSON.stringify(dxfExpr(dxf));
      if (!dxfIds.has(expr)) dxfIds.set(expr, dxfIds.size);
      return dxfIds.get(expr);
    });
    const elements = [...body.matchAll(/<tableStyleElement type="(\w+)"(?: size="(\d+)")? dxfId="(\d+)"\/>/g)]
      .map(([, type, size, id]) => {
        const index = Number(id) - base;
        if (index < 0 || index >= local.length) throw new Error(`${name}: dxfId ${id} outside a ${base}-based <dxfs> of ${local.length}`);
        return [type, Number(size ?? 1), local[index]];
      });
    if (elements.length === 0) throw new Error(`${name} has no elements`);
    out.push([name, elements]);
  }
  return out;
}
const pivotStyles = readStyles('PivotStyle', 0);
const tableStyles = readStyles('TableStyle', 1);
// Index zero represents absence, never an explicitly present empty component.
// Bounds are checked here rather than truncating values into the emitted u8/u16.
function dictionary(maxIndex = 255) {
  const values = [];
  const ids = new Map();
  return {
    values,
    intern(value) {
      if (value === null) return 0;
      const key = JSON.stringify(value);
      if (!ids.has(key)) {
        if (values.length === maxIndex) throw new Error('component dictionary exceeds index width');
        values.push(value);
        ids.set(key, values.length);
      }
      return ids.get(key);
    },
  };
}
const tints = dictionary(), colors = dictionary(), fonts = dictionary();
const patterns = dictionary(), fills = dictionary(), edgeStyles = dictionary();
const edges = dictionary(), borders = dictionary(65535), kinds = dictionary(), sizes = dictionary();
const packColor = (c) => c === null ? 0 : colors.intern([c[0], tints.intern(c[1])]);
const packedDxfs = [...dxfIds.keys()].map((expr) => {
  const [font, fill, border] = JSON.parse(expr);
  const borderId = borders.intern(border === null ? null : border.map((edge) => edges.intern(
    edge === null ? null : [edgeStyles.intern(edge[0]), packColor(edge[1])],
  )));
  return [
    fonts.intern(font === null ? null : [Number(font[0]), packColor(font[1])]),
    fills.intern(fill === null ? null : [patterns.intern(fill[0]), packColor(fill[1]), packColor(fill[2])]),
    borderId & 255, borderId >>> 8,
  ];
});
const packedElements = [];
const pack = (styles) => styles.map(([name, elements]) => {
  const start = packedElements.length;
  for (const [kind, size, dxf] of elements) {
    if (!Number.isInteger(dxf) || dxf < 0 || dxf > 65535) throw new Error('invalid dxf index');
    if (!Number.isInteger(size) || size < 0 || size > 4294967295) throw new Error('invalid band size');
    packedElements.push([kinds.intern(kind), sizes.intern(size), dxf & 255, dxf >>> 8]);
  }
  if (packedElements.length > 65535) throw new Error('element offsets exceed u16');
  return [name, start, packedElements.length];
});
const packedPivotStyles = pack(pivotStyles);
const packedTableStyles = pack(tableStyles);
let out = `// Built-in Table and PivotTable styles of ECMA-376 Part 1 Annex G
// (presetTableStyles.xml), packed without changing values or element order.
// Generated by scripts/generate-spreadsheet-style-presets.mjs; do not edit.

`;
function table(name, type, values, format = (v) => JSON.stringify(v)) {
  out += `static ${name}: [${type}; ${values.length}] = [\n`;
  for (const value of values) out += `    ${format(value)},\n`;
  out += '];\n\n';
}
table('PRESET_TINTS', 'f64', tints.values, (v) => `${v}_f64`);
table('PRESET_COLORS', '[u8; 2]', colors.values);
table('PRESET_FONTS', '[u8; 2]', fonts.values);
table('PRESET_PATTERNS', '&str', patterns.values);
table('PRESET_FILLS', '[u8; 3]', fills.values);
table('PRESET_EDGE_STYLES', '&str', edgeStyles.values);
table('PRESET_EDGES', '[u8; 2]', edges.values);
table('PRESET_BORDERS', '[u8; 6]', borders.values);
// A Dxf stores font/fill indices and a little-endian u16 border index; each
// component is decoded by the shared Rust resolver. Style element records
// similarly store kind/size indices followed by a little-endian u16 Dxf index.
table('PRESET_DXFS', '[u8; 4]', packedDxfs);
table('PRESET_KINDS', '&str', kinds.values);
table('PRESET_SIZES', 'u32', sizes.values);
table('PRESET_ELEMENTS', '[u8; 4]', packedElements);
const entry = ([name, start, end]) => `(${JSON.stringify(name)}, ${start}, ${end})`;
table('PIVOT_STYLES', 'PresetStyleEntry', packedPivotStyles, entry);
table('TABLE_STYLES', 'PresetStyleEntry', packedTableStyles, entry);
const here = dirname(fileURLToPath(import.meta.url));
writeFileSync(resolve(here, '../packages/ooxml-common/src/spreadsheet_style_presets.generated.rs'), out.trimEnd() + '\n');
console.log(`${pivotStyles.length} PivotTable + ${tableStyles.length} Table styles, ${dxfIds.size} dxfs, ${packedElements.length} elements`);
