import { describe, expect, it } from 'vitest';
import { OFFICE_FONT_SUBSTITUTIONS } from './substitutions.js';
import {
  fontSubstituteScriptClusterClass,
  fontSubstituteScriptCoversText,
  substituteEntryCoversText,
} from './substitute-script.js';

describe('script-scoped visual substitutes', () => {
  it('scopes only the Arabic visual substitutes, never metric or same-name faces', () => {
    const scoped = OFFICE_FONT_SUBSTITUTIONS
      .filter((entry) => entry.script !== undefined).map((entry) => entry.for.toLowerCase()).sort();
    expect(scoped).toEqual([
      'arabic typesetting', 'sakkal majalla', 'simplified arabic', 'traditional arabic', 'univers next arabic',
    ]);
  });

  it('distinguishes an exclusive Arabic span from a span that merely contains Arabic', () => {
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا بكم', 'exclusive')).toBe(true);
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا 2026', 'exclusive')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا 2026', 'any')).toBe(true);
    expect(fontSubstituteScriptCoversText('arabic', 'Leader 2026', 'any')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', '   ', 'exclusive')).toBe(false);
    // Invisible controls and Arabic-script digits/punctuation never prove Arabic.
    expect(fontSubstituteScriptCoversText('arabic', 'Leader\uFEFF', 'any')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', '\u061C2026', 'any')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', '\u0661\u0662\u060C', 'any')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', '\u061Cمرحبا\u200F\uFEFF', 'exclusive')).toBe(true);
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا \u0661\u0662', 'exclusive')).toBe(true);
    // Hebrew is complex script but not Arabic script.
    expect(fontSubstituteScriptCoversText('arabic', 'שלום', 'any')).toBe(false);
  });

  it('keeps Arabic combining marks (Script=Inherited) and extensions with their Arabic base', () => {
    const vocalised = '\u0645\u064E\u0631\u0652\u062D\u064E\u0628\u064B\u0627'; // مَرْحَبًا
    expect(fontSubstituteScriptCoversText('arabic', vocalised, 'exclusive')).toBe(true);
    // A lone vowel sign proves Arabic; a lone tatweel or comma does not.
    expect(fontSubstituteScriptCoversText('arabic', '\u064E', 'any')).toBe(true);
    expect(fontSubstituteScriptClusterClass('arabic', '\u064E')).toBe('proof');
    expect(fontSubstituteScriptCoversText('arabic', '\u0640', 'any')).toBe(false);
    expect(fontSubstituteScriptClusterClass('arabic', '\u0640')).toBe('extension');
    expect(fontSubstituteScriptCoversText('arabic', '\u0645\u0640\u0640\u0627\u060C', 'exclusive')).toBe(true);
    // A Latin base keeps its script even with an Arabic mark attached.
    expect(fontSubstituteScriptClusterClass('arabic', 'e\u064E')).toBe('other');
    // A generic combining mark without a base inherits (neutral).
    expect(fontSubstituteScriptClusterClass('arabic', '\u0301')).toBe('neutral');
  });

  it('lets unscoped entries cover any text', () => {
    const policy = (family: string) => OFFICE_FONT_SUBSTITUTIONS.find((entry) => entry.for === family);
    expect(substituteEntryCoversText(policy('Calibri'), 'Latin', 'exclusive')).toBe(true);
    expect(substituteEntryCoversText(policy('Sakkal Majalla'), 'Latin', 'any')).toBe(false);
  });
});
