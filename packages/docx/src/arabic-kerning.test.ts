import { describe, expect, it } from 'vitest';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, type GlyphMeasureRequest } from './layout/text.js';

function service() {
  const requests: GlyphMeasureRequest[] = [];
  const text = createTextLayoutService({
    fonts: createFontResolver([]),
    measurer: {
      fingerprint: 'arabic-kerning-fixture',
      measure: (request) => {
        requests.push(request);
        return { advancePt: [...request.text].length * request.fontSizePt / 2, ascentPt: 8, descentPt: 2 };
      },
    },
  });
  return { text, requests };
}

describe('Arabic pair kerning (WORD_ARABIC_PAIR_KERNING, #1850)', () => {
  it('kerns Arabic text whatever w:kern says', () => {
    const { text, requests } = service();
    // A run without w:kern resolves to no kerning (§17.3.2.19); Windows Word
    // still kerns its Arabic glyphs, and only those.
    const fonts = { ascii: 'Arial', highAnsi: 'Arial', eastAsia: 'Arial', complexScript: 'Arial' };
    // Word writes the controls' Latin and Arabic text as separate runs.
    for (const value of ['Z ', 'العربية لغة', ' X']) {
      text.shape({ text: value, fontSizePt: 14, kerning: false, measure: true, fonts });
    }
    const arabic = requests.filter((request) => /\p{Script=Arabic}/u.test(request.text));
    const latin = requests.filter((request) => /[ZX]/u.test(request.text) && !/\p{Script=Arabic}/u.test(request.text));
    expect(arabic.length).toBeGreaterThan(0);
    expect(arabic.every((request) => request.kerning === true)).toBe(true);
    expect(latin.length).toBeGreaterThan(0);
    expect(latin.every((request) => request.kerning === false)).toBe(true);
  });

  it('keeps w:kern for a group mixing Arabic with letters of another script', () => {
    const { text, requests } = service();
    text.shape({ text: 'AV سلام', fontSizePt: 14, kerning: false, measure: true,
      fonts: { ascii: 'Arial', highAnsi: 'Arial', eastAsia: 'Arial', complexScript: 'Arial' } });
    expect(requests.length).toBeGreaterThan(0);
    expect(requests.every((request) => request.kerning === false)).toBe(true);
  });
});
