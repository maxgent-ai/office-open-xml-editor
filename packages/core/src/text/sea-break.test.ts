import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
let {
  containsSeaScript,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
  isDictionarySeaText,
  isGraphemeFillText,
  isSeaGraphemeExtend,
  isSeaScriptCodePoint,
  seaMixedBreakOffsets,
  seaTransitionOffsets,
  seaWordBreakOffsets,
} = await import('./sea-break.js');
import { DEFAULT_KINSOKU_RULES } from './kinsoku/rules.js';

beforeEach(async () => {
  vi.resetModules();
  ({ containsSeaScript, fitSeaWordPrefix, graphemeClusterOffsets, isDictionarySeaText,
    isGraphemeFillText, isSeaGraphemeExtend, isSeaScriptCodePoint, seaMixedBreakOffsets,
    seaTransitionOffsets, seaWordBreakOffsets } = await import('./sea-break.js'));
});
afterEach(() => vi.restoreAllMocks());

function mockWords(segment: (span: string) => Iterable<{ index: number; isWordLike?: boolean }>): void {
  const Native = Intl.Segmenter;
  vi.spyOn(Intl, 'Segmenter').mockImplementation(function (locale, options) {
    return options?.granularity === 'word'
      ? { segment } as Intl.Segmenter
      : new Native(locale, options);
  });
}

describe('isSeaScriptCodePoint', () => {
  const cases: [number, boolean][] = [
    [0x0dff, false], // just below Thai
    [0x0e00, true], // Thai start
    [0x0e33, true], // THAI SARA AM
    [0x0e50, true], // THAI DIGIT ZERO
    [0x0e7f, true], // Thai end
    [0x0e80, true], // Lao start
    [0x0eff, true], // Lao end (Tibetan begins immediately at 0x0F00 — no gap)
    [0x0f00, true], // Tibetan start (grapheme-fill, #961)
    [0x0f0b, true], // TIBETAN MARK INTERSYLLABIC TSHEG
    [0x0f0d, true], // TIBETAN MARK SHAD
    [0x0fff, true], // Tibetan end
    [0x1000, true], // Myanmar start (grapheme-fill, #961)
    [0x1031, true], // MYANMAR VOWEL SIGN E
    [0x103c, true], // MYANMAR CONSONANT SIGN MEDIAL RA
    [0x109f, true], // Myanmar end
    [0x10a0, false], // just above Myanmar (Georgian)
    [0x177f, false], // just below Khmer
    [0x1780, true], // Khmer start
    [0x17d2, true], // KHMER SIGN COENG
    [0x17ff, true], // Khmer end
    [0x1800, false], // just above Khmer
    [0xa9e0, true], // Myanmar Extended-B start (grapheme-fill)
    [0xa9ff, true], // Myanmar Extended-B end
    [0xaa60, true], // Myanmar Extended-A start (grapheme-fill)
    [0xaa7f, true], // Myanmar Extended-A end
    [0xaa80, false], // just above Myanmar Extended-A
    [0x0041, false], // Latin A
    [0x3042, false], // Hiragana — CJK, not SEA
  ];
  for (const [cp, expected] of cases) {
    it(`U+${cp.toString(16).toUpperCase().padStart(4, '0')} → ${expected}`, () => {
      expect(isSeaScriptCodePoint(cp)).toBe(expected);
    });
  }
});

describe('isSeaGraphemeExtend', () => {
  const cases: [number, boolean][] = [
    [0x0e01, false], // THAI CHARACTER KO KAI (base consonant) — cluster start
    [0x0e31, true], //  THAI MAI HAN-AKAT (above vowel) — extend
    [0x0e33, true], //  THAI SARA AM — clusters onto its base (UAX#29 non-break)
    [0x0e34, true], //  THAI SARA I (above vowel) — extend
    [0x0e3a, true], //  THAI PHINTHU — extend
    [0x0e3b, false], // unassigned gap between vowel block and currency — not extend
    [0x0e40, false], // THAI SARA E (leading vowel, spacing) — its own cluster
    [0x0e48, true], //  THAI MAI EK (tone mark) — extend
    [0x0e4e, true], //  THAI YAMAKKAN — extend
    [0x0e50, false], // THAI DIGIT ZERO — cluster start
    [0x0eb1, true], //  LAO VOWEL SIGN MAI KAN — extend
    [0x0ec8, true], //  LAO TONE MAI EK — extend
    [0x17b6, true], //  KHMER VOWEL SIGN AA (spacing mark) — extend
    [0x17d2, true], //  KHMER SIGN COENG (subscript former) — extend
    [0x17dd, true], //  KHMER SIGN ATTHACAN — extend
    [0x1780, false], // KHMER LETTER KA (base) — cluster start
    [0x0041, false], // Latin A — not SEA, not extend
  ];
  for (const [cp, expected] of cases) {
    it(`U+${cp.toString(16).toUpperCase().padStart(4, '0')} → ${expected}`, () => {
      expect(isSeaGraphemeExtend(cp)).toBe(expected);
    });
  }

  // The predicate must agree with the platform grapheme segmenter for every SEA
  // code point: for two adjacent SEA code points a UAX#29 grapheme break falls
  // before the right one iff it is NOT extend (the SEA blocks have no Prepend),
  // so this is the sole cluster rule the thaiDistribute justifier relies on — and
  // unlike the segmenter it needs no `Intl.Segmenter`, staying correct in the
  // graceful-fallback runtimes.
  it('matches Intl.Segmenter grapheme clustering across all SEA code points', () => {
    if (typeof Intl?.Segmenter !== 'function') return; // environment lacks the segmenter
    const seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    const clustersOne = (s: string): boolean => [...seg.segment(s)].length === 1;
    for (const [lo, hi] of [
      [0x0e00, 0x0e7f],
      [0x0e80, 0x0eff],
      [0x1780, 0x17ff],
    ]) {
      for (let cp = lo; cp <= hi; cp++) {
        // "ก" + cp is one cluster exactly when cp extends the preceding base.
        const segmenterSaysExtend = clustersOne('ก' + String.fromCodePoint(cp));
        expect(isSeaGraphemeExtend(cp), `U+${cp.toString(16)}`).toBe(segmenterSaysExtend);
      }
    }
  });
});

describe('containsSeaScript', () => {
  it('true for Thai / Lao / Khmer text', () => {
    expect(containsSeaScript('ภาษาไทย')).toBe(true);
    expect(containsSeaScript('ພາສາລາວ')).toBe(true);
    expect(containsSeaScript('ភាសាខ្មែរ')).toBe(true);
  });
  it('true for Myanmar / Tibetan text (grapheme-fill scripts, #961)', () => {
    expect(containsSeaScript('မြန်မာဘာသာ')).toBe(true);
    expect(containsSeaScript('བོད་ཡིག')).toBe(true);
  });
  it('true when SEA is embedded in Latin', () => {
    expect(containsSeaScript('Hello ภาษา world')).toBe(true);
  });
  it('false for pure Latin / CJK / empty', () => {
    expect(containsSeaScript('Hello, world!')).toBe(false);
    expect(containsSeaScript('日本語のテキスト')).toBe(false);
    expect(containsSeaScript('')).toBe(false);
  });
});

describe('seaWordBreakOffsets — platform ICU', () => {
  it('returns interior grapheme-boundary offsets for a Thai sentence', () => {
    const text = 'ภาษาไทยเป็นภาษาที่สวยงามมาก';
    const offsets = seaWordBreakOffsets(text);
    expect(offsets.length).toBeGreaterThan(0);
    // Ascending, strictly interior (never 0 or text.length).
    for (let k = 1; k < offsets.length; k++) expect(offsets[k]).toBeGreaterThan(offsets[k - 1]);
    for (const o of offsets) {
      expect(o).toBeGreaterThan(0);
      expect(o).toBeLessThan(text.length);
    }
    // Every offset is a grapheme-cluster boundary (never mid-cluster).
    const graphemes = new Set(graphemeClusterOffsets(text));
    for (const o of offsets) expect(graphemes.has(o)).toBe(true);
    // Every offset falls between two SEA-script characters (interior to the span).
    for (const o of offsets) {
      expect(isSeaScriptCodePoint(text.codePointAt(o - 1)!)).toBe(true);
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('dictionary-breaks Lao and Khmer too', () => {
    expect(seaWordBreakOffsets('ພາສາລາວແມ່ນພາສາທີ່ສວຍງາມ').length).toBeGreaterThan(0);
    expect(seaWordBreakOffsets('ភាសាខ្មែរជាភាសាដ៏ស្រស់ស្អាត').length).toBeGreaterThan(0);
  });

  it('never adds a break touching non-SEA text (mixed Thai + Latin + digits)', () => {
    const text = 'ทดสอบ ABC 123 ภาษาไทย';
    for (const o of seaWordBreakOffsets(text)) {
      // Both sides of every returned offset must be SEA — no Latin/space/digit edge.
      expect(isSeaScriptCodePoint(text.codePointAt(o - 1)!)).toBe(true);
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('offsets are correct after an astral (surrogate-pair) non-SEA prefix (UTF-16 units)', () => {
    const emoji = '😀'; // U+1F600, 2 UTF-16 units
    const thai = 'ภาษาไทยเป็นภาษา';
    const text = emoji + thai;
    const offsets = seaWordBreakOffsets(text);
    expect(offsets.length).toBeGreaterThan(0);
    for (const o of offsets) {
      expect(o).toBeGreaterThanOrEqual(emoji.length); // never inside the emoji
      expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
    }
  });

  it('returns [] for non-SEA / empty input', () => {
    expect(seaWordBreakOffsets('Hello world')).toEqual([]);
    expect(seaWordBreakOffsets('')).toEqual([]);
  });
});

describe('seaWordBreakOffsets — grapheme-fill (Myanmar / Tibetan, #961)', () => {
  // Word (macOS) breaks these at EVERY grapheme cluster, not dictionary words nor
  // the Tibetan tsheg (sample-46 ground truth). So the offsets must be exactly the
  // interior grapheme-cluster boundaries — every one, never a dictionary subset,
  // never mid-cluster.
  const myanmar = 'မြန်မာဘာသာစကားကိုစာလုံး'; // spaceless Myanmar run
  const tibetan = 'བོད་ཡིག་ནི་ཚིག'; // tsheg-separated Tibetan run

  for (const [name, text] of [['Myanmar', myanmar], ['Tibetan', tibetan]] as const) {
    it(`${name}: offsets == interior grapheme-cluster boundaries`, () => {
      const offsets = seaWordBreakOffsets(text);
      expect(offsets.length).toBeGreaterThan(1);
      // Ascending and strictly interior.
      for (let k = 1; k < offsets.length; k++) expect(offsets[k]).toBeGreaterThan(offsets[k - 1]);
      for (const o of offsets) {
        expect(o).toBeGreaterThan(0);
        expect(o).toBeLessThan(text.length);
      }
      // EXACTLY the interior grapheme boundaries (not a dictionary subset).
      expect(offsets).toEqual(graphemeClusterOffsets(text));
      // Both sides of every break are the same-script (interior to the span).
      for (const o of offsets) {
        expect(isSeaScriptCodePoint(text.codePointAt(o - 1)!)).toBe(true);
        expect(isSeaScriptCodePoint(text.codePointAt(o)!)).toBe(true);
      }
    });
  }

  it('Myanmar break offsets do NOT collapse to the coarser ICU dictionary set', () => {
    // Property proof that this is grapheme-fill, not dictionary: there are strictly
    // more grapheme boundaries than ICU 'my' word-like boundaries for this run.
    const graphemeCount = seaWordBreakOffsets(myanmar).length;
    let wordLike = 0;
    for (const s of new Intl.Segmenter('my', { granularity: 'word' }).segment(myanmar)) {
      if (s.index > 0 && s.isWordLike) wordLike++;
    }
    expect(graphemeCount).toBeGreaterThan(wordLike);
  });

  it('keeps dictionary (Thai) and grapheme-fill (Myanmar) spans separate when adjacent', () => {
    // A Thai↔Myanmar no-space boundary must not be a break (SEA↔SEA cross-class
    // edge is left to the caller); Thai stays dictionary, Myanmar stays grapheme.
    const mixed = 'ไทยမြန်မာ'; // Thai 'ไทย' then Myanmar 'မြန်မာ', no space
    const thaiLen = 'ไทย'.length;
    for (const o of seaWordBreakOffsets(mixed)) {
      // No break exactly at the class boundary.
      expect(o).not.toBe(thaiLen);
    }
  });
});

describe('seaWordBreakOffsets — graceful fallback + filtering (mocked Intl.Segmenter)', () => {
  it('returns [] when the segmenter is unavailable', () => {
    vi.spyOn(Intl, 'Segmenter').mockImplementation(function () { throw new Error('unavailable'); });
    expect(seaWordBreakOffsets('ภาษาไทยเป็นภาษา')).toEqual([]);
  });

  it('returns [] when segmentation throws', () => {
    mockWords(() => {
      throw new Error('boom');
    });
    expect(seaWordBreakOffsets('ภาษาไทยเป็นภาษา')).toEqual([]);
  });

  it('excludes non-word (punctuation) segment starts, keeps word starts', () => {
    // Fake: split the span into fixed segments; one is punctuation (isWordLike:false).
    // span "AAAA.BBBB" (all treated as one SEA span by the test text below).
    mockWords(() => {
      // segment boundaries at 0 (word), 4 (punct '.'), 5 (word)
      return [
        { index: 0, isWordLike: true },
        { index: 4, isWordLike: false },
        { index: 5, isWordLike: true },
      ];
    });
    // Use real Thai so containsSeaScript passes; the fake decides the offsets.
    const text = 'ก'.repeat(9);
    // index 0 excluded (span-start); index 4 excluded (punctuation); index 5 kept.
    expect(seaWordBreakOffsets(text)).toEqual([5]);
  });
});

describe('seaTransitionOffsets (issue #960 — no-space SEA↔non-SEA seams)', () => {
  it('offsets each SEA↔Latin transition edge (both entering and leaving SEA)', () => {
    const text = 'เมืองBangkokคือ'; // Thai | Latin | Thai
    const bangkok = text.indexOf('B'); // Thai→Latin
    const afterBangkok = bangkok + 'Bangkok'.length; // Latin→Thai
    expect(seaTransitionOffsets(text)).toEqual([bangkok, afterBangkok]);
  });

  it('offsets Thai↔digit seams so a price like 1250 can break away from Thai', () => {
    const text = 'ราคา1250บาท'; // Thai | 1250 | Thai
    const d = text.indexOf('1');
    const afterD = d + '1250'.length;
    expect(seaTransitionOffsets(text)).toEqual([d, afterD]);
  });

  it('offsets Thai↔CJK seams', () => {
    const text = 'อาหาร寿司ราเมง'; // Thai | CJK | Thai
    const cjk = text.indexOf('寿');
    const afterCjk = text.indexOf('ร', cjk);
    expect(seaTransitionOffsets(text)).toEqual([cjk, afterCjk]);
  });

  it('never offsets a seam that touches whitespace (the space is the break)', () => {
    const text = 'ทดสอบ ABC ภาษา'; // Thai SPACE Latin SPACE Thai
    // The only SEA↔non-SEA changes are across spaces → no transition offsets.
    expect(seaTransitionOffsets(text)).toEqual([]);
  });

  it('is empty for pure SEA and for non-SEA input', () => {
    expect(seaTransitionOffsets('ภาษาไทย')).toEqual([]);
    expect(seaTransitionOffsets('Hello123')).toEqual([]);
    expect(seaTransitionOffsets('')).toEqual([]);
  });

  it('stays on UTF-16 boundaries after an astral non-SEA char', () => {
    const text = '😀ภาษา'; // emoji (2 units) → Thai
    expect(seaTransitionOffsets(text)).toEqual(['😀'.length]);
  });

  it('every offset is strictly interior and separates SEA from non-SEA', () => {
    const text = 'เมืองBangkokคือ寿司ราคา1250บาท';
    for (const o of seaTransitionOffsets(text)) {
      expect(o).toBeGreaterThan(0);
      expect(o).toBeLessThan(text.length);
      const a = isSeaScriptCodePoint(text.codePointAt(o - 1)!);
      const b = isSeaScriptCodePoint(text.codePointAt(o)!);
      expect(a).not.toBe(b); // one side SEA, the other not
    }
  });
});

describe('seaMixedBreakOffsets (issue #960 — unified dict ∪ transition ∪ CJK)', () => {
  it('equals seaWordBreakOffsets for a pure-SEA token (byte-identical wrap)', () => {
    const text = 'ภาษาไทยเป็นภาษาที่สวยงามมาก';
    expect(seaMixedBreakOffsets(text)).toEqual(seaWordBreakOffsets(text));
    expect(seaMixedBreakOffsets(text, { cjk: true })).toEqual(seaWordBreakOffsets(text));
  });

  it('adds the transition seams on top of the dictionary boundaries (Thai↔Latin)', () => {
    const text = 'เมืองBangkokคือเมืองหลวงของThailand';
    const merged = seaMixedBreakOffsets(text);
    for (const o of seaTransitionOffsets(text)) expect(merged).toContain(o);
    for (const o of seaWordBreakOffsets(text)) expect(merged).toContain(o);
    // Sorted, unique, interior.
    for (let k = 1; k < merged.length; k++) expect(merged[k]).toBeGreaterThan(merged[k - 1]);
  });

  it('adds CJK per-character opportunities only when cjk:true (mixed CJK+SEA)', () => {
    const text = '日本語のテキストとภาษาไทยが同じ';
    const withoutCjk = seaMixedBreakOffsets(text);
    const withCjk = seaMixedBreakOffsets(text, { cjk: true });
    expect(withCjk.length).toBeGreaterThan(withoutCjk.length);
    // A break BEFORE an interior ideograph (e.g. before 本 in 日本) is present.
    const between = 1; // 日|本
    expect(isSeaScriptCodePoint(text.codePointAt(0)!)).toBe(false);
    expect(withCjk).toContain(between);
    // The SEA span still breaks at its own dictionary boundaries, not mid-cluster.
    for (const o of seaWordBreakOffsets(text)) expect(withCjk).toContain(o);
  });

  it('removes kinsoku-illegal positions (no line-start-forbidden char at a head)', () => {
    // 。is line-start-forbidden: an offset BEFORE it must be dropped.
    const text = 'ภาษา。日本'; // Thai | 。 | CJK
    const dot = text.indexOf('。');
    const withKinsoku = seaMixedBreakOffsets(text, { cjk: true, kinsoku: DEFAULT_KINSOKU_RULES });
    expect(withKinsoku).not.toContain(dot); // never start a line with 。
    const withoutKinsoku = seaMixedBreakOffsets(text, { cjk: true });
    expect(withoutKinsoku).toContain(dot); // the raw seam existed before filtering
  });

  it('does NOT filter when kinsoku is disabled (§17.3.1.16 <w:kinsoku w:val="0"/>)', () => {
    const text = 'ภาษา。日本';
    const dot = text.indexOf('。');
    const disabled = { ...DEFAULT_KINSOKU_RULES, enabled: false };
    // enabled:false mirrors the CJK path's `if (!rules.enabled) return` — the
    // 。 seam survives even though 。 is in the forbidden set.
    expect(seaMixedBreakOffsets(text, { cjk: true, kinsoku: disabled })).toContain(dot);
  });

  it('keeps every offset on a grapheme-cluster boundary (no base+mark/VS/ZWJ tear)', () => {
    // 漢 + VARIATION SELECTOR-1 (U+FE00) is ONE grapheme; the CJK enumerator would
    // otherwise offer a break at index 1, inside the cluster.
    const vs = '漢︀ภาษาไทย';
    const graphemes = new Set(graphemeClusterOffsets(vs));
    for (const o of seaMixedBreakOffsets(vs, { cjk: true })) expect(graphemes.has(o)).toBe(true);
    expect(seaMixedBreakOffsets(vs, { cjk: true })).not.toContain(1); // never mid 漢︀
    // Thai base + ZWJ + Thai: the ZWJ seam (index 1) must not be offered.
    const zwj = 'ก‍า';
    for (const o of seaMixedBreakOffsets(zwj, { cjk: true })) expect(new Set(graphemeClusterOffsets(zwj)).has(o)).toBe(true);
  });

  it('never breaks beside a non-breaking space (NBSP family)', () => {
    // Thai NBSP Latin — the NBSP is non-breaking, so neither seam is offered.
    expect(seaTransitionOffsets('ภาษา Bangkok')).toEqual([]);
    expect(seaTransitionOffsets('ราคา 1250')).toEqual([]);
  });

  it('is empty for non-SEA input regardless of options', () => {
    expect(seaMixedBreakOffsets('日本語ABC', { cjk: true })).toEqual([]);
    expect(seaMixedBreakOffsets('')).toEqual([]);
  });
});

describe('fitSeaWordPrefix', () => {
  // Each char is width 10; offsets at 2,5,7 (word boundaries), text length 9.
  const text = 'abcdefghi';
  const offsets = [2, 5, 7];
  const measure = (s: string) => s.length * 10;

  it('returns the largest word boundary whose prefix fits', () => {
    // avail 55 → prefixes: [0,2)=20, [0,5)=50, [0,7)=70(overflow) → best 5.
    expect(fitSeaWordPrefix(text, offsets, 0, 55, measure)).toBe(5);
  });
  it('returns text.length when the whole remainder fits', () => {
    expect(fitSeaWordPrefix(text, offsets, 0, 999, measure)).toBe(9);
  });
  it('returns start (no progress) when even the first word overflows', () => {
    // avail 15 < first word [0,2)=20 → start.
    expect(fitSeaWordPrefix(text, offsets, 0, 15, measure)).toBe(0);
  });
  it('respects a non-zero start (continuation of a wrapped word run)', () => {
    // from start 5: boundaries >5 = [7]; avail 15 → [5,7)=20 overflow → best 5.
    expect(fitSeaWordPrefix(text, offsets, 5, 15, measure)).toBe(5);
    // avail 45 → [5,7)=20 ok, [5,9)=40 ok → 9.
    expect(fitSeaWordPrefix(text, offsets, 5, 45, measure)).toBe(9);
  });
  it('treats the final segment (past the last offset) as a whole-remainder word', () => {
    // from start 7: no offsets >7; [7,9)=20; avail 25 → 9; avail 15 → 7 (no progress).
    expect(fitSeaWordPrefix(text, offsets, 7, 25, measure)).toBe(9);
    expect(fitSeaWordPrefix(text, offsets, 7, 15, measure)).toBe(7);
  });

  it('keeps a later-fitting boundary when advance is non-monotone (negative spacing)', () => {
    // Simulate strong negative letter spacing: a longer prefix can be NARROWER
    // than a shorter one, so an early-return greedy would wrongly stop short.
    // widths: [0,2)=30, [0,5)=10, [0,7)=40, whole [0,9)=50; avail 15 → best 5.
    const w: Record<number, number> = { 2: 30, 5: 10, 7: 40, 9: 50 };
    const nonMono = (s: string) => w[s.length] ?? s.length;
    // Default (full scan) keeps the non-monotone contract.
    expect(fitSeaWordPrefix(text, offsets, 0, 15, nonMono)).toBe(5);
  });
});

describe('fitSeaWordPrefix — assumeMonotone binary-search fast path (#961)', () => {
  // For grapheme-fill scripts the offsets are DENSE (one per cluster). The
  // monotone binary search must return the SAME boundary as the full scan for any
  // monotone measure, at O(log n) measure calls instead of O(n).
  const text = 'abcdefghij'; // len 10
  const denseOffsets = [1, 2, 3, 4, 5, 6, 7, 8, 9]; // every interior boundary
  const measure = (s: string) => s.length * 10; // monotone: 10px/char

  it('matches the full scan for a monotone measure at every avail', () => {
    for (let avail = 0; avail <= 120; avail += 5) {
      const fast = fitSeaWordPrefix(text, denseOffsets, 0, avail, measure, true);
      const slow = fitSeaWordPrefix(text, denseOffsets, 0, avail, measure, false);
      expect(fast).toBe(slow);
    }
  });

  it('respects a non-zero start under binary search', () => {
    // from start 4: avail 30 → [4,7)=30 ok, [4,8)=40 no → 7.
    expect(fitSeaWordPrefix(text, denseOffsets, 4, 30, measure, true)).toBe(7);
    expect(fitSeaWordPrefix(text, denseOffsets, 4, 30, measure, false)).toBe(7);
  });

  it('returns len when the whole remainder fits, and start when nothing does', () => {
    expect(fitSeaWordPrefix(text, denseOffsets, 0, 999, measure, true)).toBe(10);
    expect(fitSeaWordPrefix(text, denseOffsets, 0, 5, measure, true)).toBe(0); // < 1 char
  });

  it('uses O(log n) measure calls, not O(n) — the perf fix', () => {
    // 1000 dense boundaries; count measure invocations. Full scan ≈ n; binary
    // search ≈ log2(n) (+ the whole-remainder probe). Assert a tight log bound.
    const big = 'x'.repeat(1000);
    const bigOffsets = Array.from({ length: 999 }, (_, k) => k + 1);
    let calls = 0;
    const counting = (s: string) => { calls++; return s.length * 10; };
    fitSeaWordPrefix(big, bigOffsets, 0, 155, counting, true); // fits ~15 chars
    expect(calls).toBeLessThan(20); // ~log2(1000)=10, plus a couple probes
  });
});

describe('isGraphemeFillText (#961)', () => {
  it('true for Myanmar / Tibetan text (drives the monotone fit)', () => {
    expect(isGraphemeFillText('မြန်မာ')).toBe(true);
    expect(isGraphemeFillText('བོད་ཡིག')).toBe(true);
    expect(isGraphemeFillText('Hello မြန်မာ')).toBe(true); // first no-space char is Myanmar
  });
  it('false for dictionary (Thai/Lao/Khmer) and non-no-space text', () => {
    expect(isGraphemeFillText('ภาษาไทย')).toBe(false); // dictionary — keep full scan
    expect(isGraphemeFillText('ພາສາລາວ')).toBe(false);
    expect(isGraphemeFillText('ភាសាខ្មែរ')).toBe(false);
    expect(isGraphemeFillText('Hello world')).toBe(false);
    expect(isGraphemeFillText('日本語')).toBe(false);
    expect(isGraphemeFillText('')).toBe(false);
  });
});

describe('isDictionarySeaText (#991)', () => {
  it('true only for pure dictionary-SEA (Thai/Lao/Khmer) content', () => {
    expect(isDictionarySeaText('ภาษาไทย')).toBe(true);
    expect(isDictionarySeaText('ພາສາລາວ')).toBe(true);
    expect(isDictionarySeaText('ភាសាខ្មែរ')).toBe(true);
    expect(isDictionarySeaText('Hello ไทย 123')).toBe(true); // non-SEA is ignored
  });
  it('false for grapheme-fill scripts, mixed SEA families, and non-SEA text', () => {
    expect(isDictionarySeaText('မြန်မာ')).toBe(false); // Myanmar
    expect(isDictionarySeaText('བོད་ཡིག')).toBe(false); // Tibetan
    // A single run mixing both families is excluded REGARDLESS of which span
    // comes first (isGraphemeFillText would classify by the first span only).
    expect(isDictionarySeaText('ไทยမြန်မာ')).toBe(false);
    expect(isDictionarySeaText('မြန်မာไทย')).toBe(false);
    expect(isDictionarySeaText('Hello world')).toBe(false);
    expect(isDictionarySeaText('日本語')).toBe(false);
    expect(isDictionarySeaText('')).toBe(false);
  });
});

describe('graphemeClusterOffsets', () => {
  it('keeps a base + combining/tone mark together (Thai SARA AM cluster)', () => {
    // U+0E33 SARA AM is a single code point but base+mark cluster candidates like
    // นํ (NO NU + NIKHAHIT) must not split. Use กำ = ก + SARA AM already atomic;
    // test the classic ก + tone mark ่ (U+0E48) cluster instead.
    const text = 'ก่ข้'; // ก+MAI EK, ข+MAI THO → two clusters
    const offsets = graphemeClusterOffsets(text);
    // Exactly one interior boundary, between the two clusters (after index 2).
    expect(offsets).toEqual([2]);
  });
  it('matches the extended-grapheme segmenter for Latin-1 text, including CR LF and controls', () => {
    const native = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    const expected = (text: string) => [...native.segment(text)].map((part) => part.index).filter((index) => index > 0);
    // Every ordered pair of Basic Latin / Latin-1 Supplement code units covers
    // each rule that can apply inside the range (GB3 CR × LF, GB4/GB5 around
    // controls and the soft hyphen, GB999 elsewhere).
    for (let first = 0; first <= 0xff; first += 1) {
      for (let second = 0; second <= 0xff; second += 1) {
        const text = String.fromCharCode(first, second);
        expect(graphemeClusterOffsets(text)).toEqual(expected(text));
      }
    }
    for (const text of ['', 'a', 'Caf\u00e9 cr\u00e8me', 'one\r\ntwo\r\n\r\n', '\r\r\n\n', 'x\u00ad\ty']) {
      expect(graphemeClusterOffsets(text)).toEqual(expected(text));
    }
  });
  it('matches the segmenter across the CJK and kana fast-path ranges', () => {
    const native = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    const expected = (text: string) => [...native.segment(text)].map((part) => part.index).filter((index) => index > 0);
    const ranges = [[0x3000, 0x303f], [0x3041, 0x30ff], [0x3400, 0x4dbf], [0x4e00, 0x9fff], [0xff01, 0xff9f]];
    for (const [first, last] of ranges) {
      for (let code = first; code <= last; code += 1) {
        // `a c a c c` exposes every UAX #29 rule that could join c to a
        // neighbour: Extend/ZWJ/SpacingMark (a c), Prepend (c a), and Hangul
        // jamo or Regional_Indicator pairs (c c). Characters outside the fast
        // path (U+302A..U+302F, U+3099..U+309A, U+FF9E..U+FF9F) take the
        // segmenter and must still match it.
        const c = String.fromCharCode(code);
        const text = `a${c}a${c}${c}`;
        expect(graphemeClusterOffsets(text)).toEqual(expected(text));
      }
    }
    for (const text of ['\u65e5\u672c\u8a9e\u3002', '\u304b\u3099\u304d', '\uff76\uff9e\uff77\uff9e', '\u6f22\r\n\u304b', '\u301c\u300c\u5f15\u300d']) {
      expect(graphemeClusterOffsets(text)).toEqual(expected(text));
    }
  });
  it('returns interior offsets only (excludes 0 and length)', () => {
    const offsets = graphemeClusterOffsets('abc');
    for (const o of offsets) {
      expect(o).toBeGreaterThan(0);
      expect(o).toBeLessThan(3);
    }
  });
});
