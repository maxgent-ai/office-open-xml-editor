import { describe, expect, it } from 'vitest';
import { buildFont } from './renderer.js';

const context = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };

// These ASCII CSS strings are also JSON strings. Decode rather than compare
// an implementation's escaping algorithm: the requested names must survive.
function leadingFamilies(font: string, count: number): string[] {
  const strings = font.match(/"(?:[^"\\]|\\.)*"/gu) ?? [];
  return strings.slice(0, count).map((token) => JSON.parse(token) as string);
}

describe('DrawingML font family CSS serialization', () => {
  it('preserves quotes in the primary authored family as one CSS string', () => {
    expect(leadingFamilies(buildFont(false, false, 20, 'Face "Quoted"', context), 1))
      .toEqual(['Face "Quoted"']);
  });

  it('preserves a literal backslash and inserts complete application defaults after an escaped primary', () => {
    const families = ['Face \\31 "Quoted"', 'Default "One", Two', 'Default \\2'];
    const font = buildFont(false, false, 20, families[0], context, '', true, families.slice(1));
    expect(leadingFamilies(font, 3)).toEqual(families);
  });

  it('keeps replacement tokens in default family names literal', () => {
    const font = buildFont(false, false, 20, 'Primary', context, '', true, ['Default $&']);
    expect(leadingFamilies(font, 2)).toEqual(['Primary', 'Default $&']);
  });
});
