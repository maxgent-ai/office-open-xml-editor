import { describe, expect, it } from 'vitest';
import {
  OFFICE_FONT_SUBSTITUTIONS,
  loadedRegularAliases,
  matchFontDeclaration,
  substitutionsFor,
  type LoadedFaceFacts,
} from './substitutions.js';

describe('vendor-neutral Office substitutions', () => {
  // The served family and scope of every former table entry are pinned by the
  // frozen oracle in presets/google-fonts.test.ts; the policy itself names
  // no provider and no URL.
  it('carry no URL and no provider naming', () => {
    expect(JSON.stringify(OFFICE_FONT_SUBSTITUTIONS)).not.toMatch(/https?:|google/i);
    expect(Object.isFrozen(OFFICE_FONT_SUBSTITUTIONS)).toBe(true);
  });
});

describe('declaration matching', () => {
  it('prefers the authored family, then a declared substitute, case- and quote-insensitively', () => {
    expect(matchFontDeclaration('"Calibri"', 'calibri', OFFICE_FONT_SUBSTITUTIONS)).toEqual({ kind: 'family' });
    expect(matchFontDeclaration('carlito', ' Calibri ', OFFICE_FONT_SUBSTITUTIONS))
      .toEqual({ kind: 'substitute', substitution: { for: 'Calibri', use: 'Carlito' } });
    expect(matchFontDeclaration('Noto Naskh Arabic', 'Sakkal Majalla', OFFICE_FONT_SUBSTITUTIONS))
      .toEqual({ kind: 'substitute', substitution: { for: 'Sakkal Majalla', use: 'Noto Naskh Arabic', script: 'arabic' } });
  });

  it('never matches an undeclared pairing or a substitute in the reverse direction', () => {
    expect(matchFontDeclaration('Carlito', 'Calibri', undefined)).toBeNull();
    expect(matchFontDeclaration('Calibri', 'Carlito', OFFICE_FONT_SUBSTITUTIONS)).toBeNull();
    expect(matchFontDeclaration('Carlito', 'Calibri Light', OFFICE_FONT_SUBSTITUTIONS)).toBeNull();
    expect(substitutionsFor(OFFICE_FONT_SUBSTITUTIONS, 'Cambria Math')).toEqual([]);
  });
});

describe('loadedRegularAliases', () => {
  const face = (family: string, patch: Partial<LoadedFaceFacts> = {}): LoadedFaceFacts =>
    ({ family, weight: '400', style: 'normal', status: 'loaded', ...patch });

  it('offers a full regular name only for a loaded weight-400 upright face of any source', () => {
    const aliases = loadedRegularAliases([
      face('"Lato"'),
      face('Roboto', { status: 'error' }),
      face('Raleway', { weight: '700' }),
      face('Poppins', { style: 'italic' }),
      face('Carlito', { weight: 'normal' }),
    ]);
    expect([...aliases]).toEqual([['lato regular', 'Lato'], ['carlito regular', 'Carlito']]);
  });
});
