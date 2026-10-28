import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parityGroups } from './test-support/word-space-fit-parity.test-support.js';

// Partition/width digests for excluded mixed-script groups retain origin/main
// 4a387ebcb. Latin compressPunctuation groups below mode 15 incorporate the
// kerning, OpenType (mode 14) and lineWrapLikeWord6 (mode 11) independence
// confirmed by Windows Word controls; their xAvg fit
// and rejection boundaries are independently tested in reference-font-layout.
// Other Latin settings retain the historical results.
const EXPECTED = JSON.parse(readFileSync(
  new URL('./word-space-fit-parity-main.json', import.meta.url), 'utf8',
)) as Record<string, string>;

describe('Word space-fit scope regression', () => {
  it('preserves excluded mixed settings and the corrected Latin fit policy', () => {
    const groups = Object.fromEntries(parityGroups());
    expect(Object.keys(groups).length).toBe(Object.keys(EXPECTED).length);
    expect(groups).toEqual(EXPECTED);
  }, 60_000);
});
