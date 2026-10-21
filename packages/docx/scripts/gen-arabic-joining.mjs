// Generates packages/docx/src/arabic-joining.generated.ts from the Unicode
// Character Database (UCD). Run:
//   node packages/docx/scripts/gen-arabic-joining.mjs
//
// Emits the ArabicShaping.txt Joining_Group families used by Word kashida
// justification (ECMA-376 §17.18.44 jc=lowKashida/mediumKashida/highKashida).
// The Joining_Type table itself is shared text data generated in core
// (packages/core/scripts/gen-joining-type.mjs).
//
// The generated file is data straight from the UCD — never hand-edit it.

import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const VERSION = '17.0.0';
const BASE = `https://www.unicode.org/Public/${VERSION}/ucd`;
const ARABIC_SHAPING_SRC = 'ArabicShaping.txt';
const HERE = dirname(fileURLToPath(import.meta.url));
const ARABIC_SHAPING_LOCAL = join(HERE, ARABIC_SHAPING_SRC);
const OUT = join(HERE, '..', 'src', 'arabic-joining.generated.ts');

async function readArabicShapingSource() {
  try {
    return await readFile(ARABIC_SHAPING_LOCAL, 'utf8');
  } catch {
    const url = `${BASE}/${ARABIC_SHAPING_SRC}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`GET ${url} -> ${res.status}`);
    return res.text();
  }
}

function fmtArray(nums, perLine = 16) {
  const out = [];
  for (let i = 0; i < nums.length; i += perLine) {
    out.push('  ' + nums.slice(i, i + perLine).join(', ') + ',');
  }
  return out.join('\n');
}

function fmtHexArray(nums, perLine = 12) {
  return fmtArray(nums.map((cp) => `0x${cp.toString(16).toUpperCase()}`), perLine);
}

/*
 * UCD Joining_Group -> SCRIPT_JUSTIFY family mapping from old HarfBuzz's
 * harfbuzz-arabic.c getArabicProperties: Seen/Sad -> Seen; Hah -> HaaDal;
 * Beh followed by Reh/Yeh -> BaRa; Alef/Tah (with Kaf/Gaf/Lam aliases) ->
 * Alef; Waw/Ain (with Feh/Qaf aliases) -> Waw.
 */
const KASHIDA_GROUP_SPECS = [
  ['KASHIDA_SEEN_PREV', new Set(['SEEN', 'SAD']), 17],
  ['KASHIDA_HAH_PREV', new Set(['HAH']), 22],
  ['KASHIDA_BEH_PREV', new Set(['BEH']), 27],
  ['KASHIDA_REH_YEH_CUR', new Set(['REH', 'YEH', 'FARSI YEH']), 37],
  ['KASHIDA_ALEF_TAH_CUR', new Set(['ALEF', 'TAH', 'KAF', 'GAF', 'LAM']), 68],
  ['KASHIDA_WAW_AIN_CUR', new Set(['WAW', 'AIN', 'FEH', 'QAF']), 41],
];

function buildKashidaGroups(text) {
  const values = new Map(KASHIDA_GROUP_SPECS.map(([name]) => [name, new Set()]));

  for (const raw of text.split('\n')) {
    const line = raw.split('#')[0].trim();
    if (!line) continue;
    const fields = line.split(';').map((field) => field.trim());
    if (fields.length < 4) continue;

    const range = fields[0].match(/^([0-9A-Fa-f]+)(?:\.\.([0-9A-Fa-f]+))?$/);
    if (!range) continue;
    const start = parseInt(range[1], 16);
    const end = range[2] ? parseInt(range[2], 16) : start;
    const joiningGroup = fields[3];

    for (const [name, acceptedGroups] of KASHIDA_GROUP_SPECS) {
      if (!acceptedGroups.has(joiningGroup)) continue;
      const set = values.get(name);
      for (let cp = start; cp <= end; cp++) set.add(cp);
    }
  }

  return Object.fromEntries(KASHIDA_GROUP_SPECS.map(([name, , expectedCount]) => {
    const result = [...values.get(name)].sort((a, b) => a - b);
    if (result.length === 0 || result.length !== expectedCount) {
      throw new Error(
        `${name}: expected ${expectedCount} Joining_Group code points, got ${result.length}`,
      );
    }
    return [name, result];
  }));
}

async function main() {
  const arabicShapingText = await readArabicShapingSource();
  const kashidaGroups = buildKashidaGroups(arabicShapingText);
  const arabicShapingVer =
    arabicShapingText.match(/#\s*ArabicShaping-([\d.]+)\.txt/)?.[1] ?? VERSION;
  const arabicShapingDate =
    arabicShapingText.match(/#\s*Date:\s*([0-9-]+)/)?.[1] ?? 'unknown';

  const body = `// AUTO-GENERATED from the Unicode Character Database (UCD ${arabicShapingVer}).
// Source: ${BASE}/${ARABIC_SHAPING_SRC} (ArabicShaping-${arabicShapingVer}.txt, dated ${arabicShapingDate}).
// Joining_Group families follow harfbuzz-arabic.c getArabicProperties' group-to-class mapping.
// DO NOT EDIT — regenerate via packages/docx/scripts/gen-arabic-joining.mjs
/* eslint-disable */

${Object.entries(kashidaGroups).map(([name, codePoints]) => `export const ${name}: number[] = [
${fmtHexArray(codePoints)}
];`).join('\n\n')}
`;

  await writeFile(OUT, body, 'utf8');
  console.log(`wrote ${OUT}\n  UCD version: ${arabicShapingVer} (${arabicShapingDate})`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
