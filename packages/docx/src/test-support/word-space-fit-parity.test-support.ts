// Inputs outside WORD_COMPRESSED_SPACE_LINE_FIT whose line partitions and
// segment widths must equal origin/main (4a387ebcb): Latin-only paragraphs in
// every compatibility setting, and mixed East Asian / Latin paragraphs in
// every setting the rule excludes.
import { createHash } from 'node:crypto';
import { layoutStubParagraph, type StubParagraph, type StubRun } from './word-space-fit.test-support.js';

const LATIN: readonly StubRun[] = [
  { text: 'AV To + AV To + To AV', ascii: 'Arial', eastAsia: 'Arial', sizePt: 8.5, bold: true },
  { text: 'nnnn nnnn nn n nnnn nn', ascii: 'Arial', eastAsia: 'Arial', sizePt: 16, bold: false },
  { text: 'ABCD AB + 1 ABCD  AB', ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true },
];
const MIXED: readonly StubRun[] = [
  { text: '甲甲甲甲 + 乙乙 + 丙丙丙丙', ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true },
  { text: '甲AV甲 + 乙乙 + 丙To丙丙', ascii: 'Arial', eastAsia: 'Meiryo', sizePt: 8.5, bold: true },
];
const SETTINGS = {
  modes: [undefined, 12, 14, 15],
  spacing: [undefined, 'compressPunctuation', 'compressPunctuationAndJapaneseKana', 'doNotCompress'],
} as const;
const OFFSETS_PT = [-6, -3, -2, -1, -0.5, 0.25];

function natural(run: StubRun): number {
  return layoutStubParagraph({
    runs: [run], environment: {}, bandPt: 10_000, justification: 'left',
  })[0]!.reduce((sum, segment) => sum + segment.width, 0);
}

/** Excluded mixed settings: everything except an authored mode below 15 with a
 * compressing characterSpacingControl and Word 6 wrapping off. */
function mixedExcluded(mode: number | undefined, spacing: string | undefined, word6: boolean) {
  return word6 || mode === undefined || mode >= 15
    || (spacing !== 'compressPunctuation' && spacing !== 'compressPunctuationAndJapaneseKana');
}

export function parityGroups(): Map<string, string> {
  const groups = new Map<string, string>();
  const cases: [string, StubParagraph][] = [];
  for (const [kind, runs] of [['latin', LATIN], ['mixed', MIXED]] as const) {
    for (const run of runs) {
      const width = natural(run);
      for (const mode of SETTINGS.modes) {
        for (const spacing of SETTINGS.spacing) {
          for (const word6 of [false, true]) {
            if (kind === 'mixed' && !mixedExcluded(mode, spacing, word6)) continue;
            for (const openType of [false, true]) {
              for (const kerning of [undefined, 0]) {
                for (const justification of ['left', 'both', 'distribute'] as const) {
                  for (const offset of OFFSETS_PT) {
                    cases.push([`${kind}|${run.text}|${mode}|${spacing}|${word6}`, {
                      runs: [{ ...run, kerning }],
                      environment: {
                        ...(mode === undefined ? {} : { compatibilityMode: mode }),
                        ...(spacing === undefined ? {} : { characterSpacingControl: spacing }),
                        enableOpenTypeFeatures: openType,
                        lineWrapLikeWord6: word6,
                      },
                      bandPt: width + offset,
                      justification,
                    }]);
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  const results = new Map<string, unknown[]>();
  for (const [key, paragraph] of cases) {
    const lines = layoutStubParagraph(paragraph)
      .map((line) => line.map(({ text, width }) => ({ text, width })));
    results.set(key, [...(results.get(key) ?? []), lines]);
  }
  for (const [key, value] of results) {
    groups.set(key, createHash('sha256').update(JSON.stringify(value)).digest('hex').slice(0, 16));
  }
  return groups;
}
