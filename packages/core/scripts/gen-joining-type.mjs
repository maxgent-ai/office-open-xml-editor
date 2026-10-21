// Generates packages/core/src/text/joining-type.generated.ts from the Unicode
// Character Database (UCD). Run:
//   node packages/core/scripts/gen-joining-type.mjs
//
// Emits, as a gap-free run-length table, the Joining_Type (jt) property for
// every code point in [0, 0x110000) from extracted/DerivedJoiningType.txt,
// including its `@missing` block default (Non_Joining).
//
// Joining_Type (UAX #44 / ArabicShaping.txt) drives cursive joining:
//   U  Non_Joining   — does not join either side (default, the `@missing` value).
//   C  Join_Causing  — forces a join on both sides (U+0640 TATWEEL, ZWJ).
//   D  Dual_Joining  — joins on both sides (most Arabic letters: beh, seen, …).
//   L  Left_Joining  — joins only to the following letter (rare).
//   R  Right_Joining — joins only to the preceding letter (alef, dal, reh, waw, …).
//   T  Transparent   — combining marks, skipped when testing whether two
//                      letters join across them.
// Consumers: DOCX kashida justification and the WMF text player.
//
// The generated file is data straight from the UCD — never hand-edit it.

import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const VERSION = '17.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd`;
const SRC = 'DerivedJoiningType.txt';
const SRC_PATH = `extracted/${SRC}`;
const HERE = dirname(fileURLToPath(import.meta.url));
const LOCAL = join(HERE, SRC);
const OUT = join(HERE, '..', 'src', 'text', 'joining-type.generated.ts');
const MAX_CP = 0x110000;

// Canonical jt value index order. Index is what we store; names map back in
// arabic-joining.ts. Both the short single-letter codes used in the data lines
// (C/D/L/R/T) and the long `@missing` names map to these.
const JT_NAMES = ['U', 'C', 'D', 'L', 'R', 'T'];
const JT_INDEX = Object.fromEntries(JT_NAMES.map((n, i) => [n, i]));
const LONG_TO_SHORT = {
  Non_Joining: 'U',
  Join_Causing: 'C',
  Dual_Joining: 'D',
  Left_Joining: 'L',
  Right_Joining: 'R',
  Transparent: 'T',
};

// Prefer the checked-in local copy (so the generator is reproducible offline and
// records the exact bytes we shipped); fall back to fetching the canonical URL.
async function readSource() {
  try {
    return await readFile(LOCAL, 'utf8');
  } catch {
    const url = `${BASE}/${SRC_PATH}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
    return res.text();
  }
}

function jtIndex(name) {
  const short = LONG_TO_SHORT[name] ?? name;
  if (!(short in JT_INDEX)) throw new Error(`unknown Joining_Type: ${name}`);
  return JT_INDEX[short];
}

function buildTable(text) {
  // Default everything to the file-wide @missing value (Non_Joining), then
  // overlay explicit assigned ranges. Result covers [0, MAX_CP) with no gaps.
  const jt = new Uint8Array(MAX_CP).fill(JT_INDEX.U);

  const apply = (start, end, idx) => {
    for (let cp = start; cp <= end && cp < MAX_CP; cp++) jt[cp] = idx;
  };

  // Honour any `@missing` block defaults in file order (there is one: Non_Joining).
  for (const raw of text.split('\n')) {
    const missing = raw.match(
      /^#\s*@missing:\s*([0-9A-Fa-f]+)\.\.([0-9A-Fa-f]+)\s*;\s*(\w+)/,
    );
    if (missing) {
      apply(parseInt(missing[1], 16), parseInt(missing[2], 16), jtIndex(missing[3]));
    }
  }
  // Overlay explicit ranges. Field 1 is one of C / D / L / R / T.
  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const m = line.match(/^([0-9A-Fa-f]+)(?:\.\.([0-9A-Fa-f]+))?\s*;\s*(\w+)$/);
    if (!m) continue;
    const start = parseInt(m[1], 16);
    const end = m[2] ? parseInt(m[2], 16) : start;
    apply(start, end, jtIndex(m[3]));
  }

  // Run-length compress into gap-free ranges: range i covers [starts[i], starts[i+1]).
  const starts = [0];
  const values = [jt[0]];
  for (let cp = 1; cp < MAX_CP; cp++) {
    if (jt[cp] !== values[values.length - 1]) {
      starts.push(cp);
      values.push(jt[cp]);
    }
  }
  return { starts, values };
}

function fmtArray(nums, perLine = 16) {
  const out = [];
  for (let i = 0; i < nums.length; i += perLine) {
    out.push('  ' + nums.slice(i, i + perLine).join(', ') + ',');
  }
  return out.join('\n');
}

async function main() {
  const text = await readSource();
  // Record the exact UCD file version from its header for the provenance comment.
  const headerVer = text.match(/#\s*DerivedJoiningType-([\d.]+)\.txt/)?.[1] ?? VERSION;
  const headerDate = text.match(/#\s*Date:\s*([0-9-]+)/)?.[1] ?? 'unknown';
  const { starts, values } = buildTable(text);

  const body = `// AUTO-GENERATED from the Unicode Character Database (UCD ${headerVer}).
// Source: ${BASE}/${SRC_PATH} (DerivedJoiningType-${headerVer}.txt, dated ${headerDate}).
// Property: Joining_Type (jt), UAX #44 / ArabicShaping.txt.
// DO NOT EDIT — regenerate via packages/core/scripts/gen-joining-type.mjs
/* eslint-disable */

export const UNICODE_VERSION = '${headerVer}';

/** Canonical Joining_Type value order (index into these tables). */
export const JT_NAMES = [
  ${JT_NAMES.map((n) => `'${n}'`).join(', ')},
] as const;

/** Range starts (sorted, gap-free): range i covers [JT_RANGE_STARTS[i], JT_RANGE_STARTS[i+1]). */
export const JT_RANGE_STARTS: number[] = [
${fmtArray(starts)}
];

/** Joining_Type index (into JT_NAMES) for the range beginning at the matching JT_RANGE_STARTS entry. */
export const JT_RANGE_VALUE: number[] = [
${fmtArray(values)}
];
`;

  await writeFile(OUT, body, 'utf8');
  console.log(
    `wrote ${OUT}\n  UCD version: ${headerVer} (${headerDate})\n  jt ranges: ${starts.length}`,
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
