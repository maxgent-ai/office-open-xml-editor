import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parityGroups } from './test-support/word-space-fit-parity.test-support.js';

// Outside WORD_COMPRESSED_SPACE_LINE_FIT the line breaker must be unchanged:
// `word-space-fit-parity-main.json` holds digests of every partition and
// segment width produced by origin/main 4a387ebcb for the same inputs
// (Latin-only paragraphs in all compatibility settings; mixed paragraphs in
// every excluded setting; OpenType, kerning and alignment variants).
const MAIN = JSON.parse(readFileSync(
  new URL('./word-space-fit-parity-main.json', import.meta.url), 'utf8',
)) as Record<string, string>;

describe('WORD_COMPRESSED_SPACE_LINE_FIT scope parity with main', () => {
  it('reproduces main for every out-of-scope input group', () => {
    const groups = Object.fromEntries(parityGroups());
    expect(Object.keys(groups).length).toBe(Object.keys(MAIN).length);
    expect(groups).toEqual(MAIN);
  }, 60_000);
});
