// Line breaking for scripts written WITHOUT inter-word spaces, where a line may
// break between two characters that have no whitespace between them. Two families
// need different opportunity sources, both funnelled through the same wrap kernel
// ({@link fitSeaWordPrefix}) so every renderer's wrap loop consumes one contract:
//
//   • DICTIONARY scripts — Thai U+0E00–0E7F, Lao U+0E80–0EFF, Khmer U+1780–17FF
//     (GitHub #797 / #955). Word boundaries here are lexical and cannot be
//     derived from Unicode ranges; they require a dictionary. We enumerate them
//     with the platform ICU dictionary via `Intl.Segmenter({granularity:'word'})`.
//
//   • GRAPHEME-FILL scripts — Myanmar U+1000–109F (+ Extended-A U+AA60–AA7F,
//     Extended-B U+A9E0–A9FF) and Tibetan U+0F00–0FFF (GitHub #961). Word (macOS
//     layout) does NOT break these at dictionary words nor at the Tibetan tsheg
//     `་`: ground-truth measurement (Word export of sample-46) shows it fills each
//     line to the column edge and breaks at ANY grapheme-cluster boundary — 4 of 6
//     observed Tibetan breaks and 1 of 5 Myanmar breaks split a syllable/word
//     mid-cluster (e.g. the Myanmar word `သည်` broken `သ`|`ည်`), and it ships a
//     Burmese dictionary it declines to use. So the break opportunities for these
//     scripts are EVERY interior grapheme-cluster boundary; the sole invariant is
//     "never split a base + stacked/combining cluster". Routing them here — rather
//     than the renderers' generic over-long-word splitter, which splits at
//     CODE-POINT granularity and can tear a cluster (Myanmar `တို`→`တိ`|`ု`) —
//     upgrades that split to grapheme granularity, matching Word.
//
// Design invariants (shared with the docx/pptx/xlsx breakers):
//   • ADD break opportunities only; never change glyphs or advances. The run is
//     split only at the ACTUAL line break, so within a line the text stays one
//     contiguous draw (measure==paint).
//   • Restrict added opportunities to boundaries INTERIOR to a maximal same-class
//     no-space span — between two dictionary WORDS (dictionary scripts) or between
//     two grapheme CLUSTERS (grapheme-fill scripts). A boundary that touches
//     non-SEA text (Latin, digits, CJK, whitespace) is left to the existing
//     whitespace/Latin/CJK logic, so non-SEA content stays byte-identical.
//   • Graceful fallback: when `Intl.Segmenter` is unavailable (old runtimes, a
//     `small-icu` build without the dictionaries) or throws, dictionary scripts
//     return no offsets and grapheme-fill scripts fall back to code-point
//     boundaries; the caller keeps its current cluster/character behaviour.
//
// The no-space SEA↔Latin/CJK/digit script-transition boundary (#960) is added by
// {@link seaTransitionOffsets}; {@link seaMixedBreakOffsets} unions the dictionary
// (or grapheme-fill) word/cluster boundaries, the transition boundaries and the
// CJK per-character opportunities for a single run that mixes these scripts with
// Latin/digits/CJK.

import { isCjkBreakChar } from './cjk-ranges.js';

/** No-inter-word-space script tag. Dictionary scripts (`th`/`lo`/`km`) pick the
 *  ICU dictionary — ICU dispatches by the character's SCRIPT, not this locale, so
 *  it is only a hint. Grapheme-fill scripts (`my`/`bo`, #961) take no dictionary:
 *  every grapheme-cluster boundary is a break opportunity. */
export type SeaScript = 'th' | 'lo' | 'km' | 'my' | 'bo';

/** Grapheme-fill scripts (Myanmar, Tibetan): break at every interior grapheme
 *  cluster, never a dictionary word (Word ground truth, #961). */
function isGraphemeFillScript(s: SeaScript): boolean {
  return s === 'my' || s === 'bo';
}

/**
 * True when `cp` belongs to a no-inter-word-space script — the dictionary scripts
 * Thai/Lao/Khmer or the grapheme-fill scripts Myanmar/Tibetan. This is a SCRIPT-
 * membership test, not a break predicate: the blocks include combining marks,
 * tone/stacked marks, digits and punctuation that are NOT independently
 * breakable — the actual break opportunities come from {@link seaWordBreakOffsets}.
 * Used to detect spans and as the cheap gate in {@link containsSeaScript}.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isSeaScriptCodePoint(cp: number): boolean {
  return (
    (cp >= 0x0e00 && cp <= 0x0e7f) || // Thai (dictionary)
    (cp >= 0x0e80 && cp <= 0x0eff) || // Lao (dictionary)
    (cp >= 0x1780 && cp <= 0x17ff) || // Khmer (dictionary)
    (cp >= 0x1000 && cp <= 0x109f) || // Myanmar (grapheme-fill)
    (cp >= 0xaa60 && cp <= 0xaa7f) || // Myanmar Extended-A (grapheme-fill)
    (cp >= 0xa9e0 && cp <= 0xa9ff) || // Myanmar Extended-B (grapheme-fill)
    (cp >= 0x0f00 && cp <= 0x0fff) // Tibetan (grapheme-fill)
  );
}

/**
 * True when `cp` is a Southeast-Asian (Thai/Lao/Khmer) combining mark that
 * EXTENDS the preceding grapheme cluster rather than starting a new one — i.e.
 * its UAX#29 Grapheme_Cluster_Break is `Extend` or `SpacingMark` (above/below
 * vowel signs, tone marks, the Khmer COENG subscript-former, etc.). These are
 * exactly the code points a `thaiDistribute` justifier must NOT open a gap
 * before, so a base consonant keeps its marks glued.
 *
 * For a boundary between two SEA code points this predicate is sufficient to
 * decide grapheme clustering: the SEA blocks contain no `Prepend` characters
 * (verified against UCD), so a break falls before every SEA code point EXCEPT an
 * Extend/SpacingMark. Ranges derived from the Unicode Character Database
 * (Grapheme_Cluster_Break), cross-checked against `Intl.Segmenter` grapheme
 * output — so this stays correct even where the platform `Intl.Segmenter` is
 * unavailable (the {@link graphemeClusterOffsets} code-point fallback is NOT
 * cluster-safe and must not gate SEA distribution).
 */
export function isSeaGraphemeExtend(cp: number): boolean {
  return (
    // Thai (U+0E00–0E7F): MAI HAN-AKAT; SARA AM..PHINTHU; MAITAIKHU..YAMAKKAN.
    cp === 0x0e31 ||
    (cp >= 0x0e33 && cp <= 0x0e3a) ||
    (cp >= 0x0e47 && cp <= 0x0e4e) ||
    // Lao (U+0E80–0EFF): MAI KAN; SIGN PALI VIRAMA..semivowel signs; tone marks.
    cp === 0x0eb1 ||
    (cp >= 0x0eb3 && cp <= 0x0ebc) ||
    (cp >= 0x0ec8 && cp <= 0x0ece) ||
    // Khmer (U+1780–17FF): vowel signs, COENG, robat, signs (17B4–17D3) + ATTHACAN.
    (cp >= 0x17b4 && cp <= 0x17d3) ||
    cp === 0x17dd
  );
}

/** The no-space script of a code point already known to be one (see
 *  {@link isSeaScriptCodePoint}). Dictionary scripts pick the per-span ICU
 *  dictionary; grapheme-fill scripts pick the grapheme path. */
function seaScriptOf(cp: number): SeaScript {
  if (cp <= 0x0e7f) return 'th'; // 0E00–0E7F Thai
  if (cp <= 0x0eff) return 'lo'; // 0E80–0EFF Lao
  if (cp <= 0x0fff) return 'bo'; // 0F00–0FFF Tibetan
  if (cp <= 0x109f) return 'my'; // 1000–109F Myanmar
  if (cp <= 0x17ff) return 'km'; // 1780–17FF Khmer
  return 'my'; // AA60–AA7F / A9E0–A9FF Myanmar Extended-A/B
}

/**
 * Cheap presence gate: true when any code point of `text` is a no-inter-word-space
 * script (Thai/Lao/Khmer or Myanmar/Tibetan). Used by the renderers to skip all
 * segmentation work for the overwhelmingly common non-SEA input (zero
 * `Intl.Segmenter` cost), mirroring docx's `hasCJKBreakOpportunity`.
 */
export function containsSeaScript(text: string): boolean {
  for (const ch of text) {
    if (isSeaScriptCodePoint(ch.codePointAt(0)!)) return true;
  }
  return false;
}

/**
 * True when `text`'s no-space script is grapheme-fill (Myanmar/Tibetan) rather
 * than dictionary (Thai/Lao/Khmer) — decided by its FIRST no-space character
 * (a segment produced by the renderers is a single same-class span). Callers pass
 * this as {@link fitSeaWordPrefix}'s `assumeMonotone` so a grapheme-fill run — whose
 * break offsets are dense (one per cluster) — takes the O(log n) binary-search fit
 * instead of the O(n) full scan the dictionary path needs. Returns false for pure
 * dictionary text (keep the safe full scan) and for non-no-space text.
 */
export function isGraphemeFillText(text: string): boolean {
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (isSeaScriptCodePoint(cp)) return isGraphemeFillScript(seaScriptOf(cp));
  }
  return false;
}

/**
 * True when `text` carries dictionary-SEA (Thai/Lao/Khmer) content and NO
 * grapheme-fill (Myanmar/Tibetan) content. Unlike {@link isGraphemeFillText}
 * this scans EVERY code point, so a rare single run mixing both SEA families
 * is excluded rather than classified by its first span.
 *
 * Introduced for the issue #991 Word-verified fit rules (zero trailing-space
 * shrink on Thai lines; whole-chunk movement of a full-line-fitting no-space
 * run): they apply exactly to pure dictionary-SEA segments. Grapheme-fill
 * scripts keep the per-cluster greedy fill (#961), and a mixed-family segment
 * falls back to the pre-#991 greedy path so a grapheme-fill span is never
 * moved as part of an atomic chunk.
 */
export function isDictionarySeaText(text: string): boolean {
  let hasDict = false;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (!isSeaScriptCodePoint(cp)) continue;
    if (isGraphemeFillScript(seaScriptOf(cp))) return false;
    hasDict = true;
  }
  return hasDict;
}

// ── Segmenter plumbing (cached) ────────────────────────────

/** Minimal shape of one `Intl.Segments` entry we rely on. */
interface SeaSegment {
  index: number;
  isWordLike?: boolean;
}
const intlWordCache = new Map<SeaScript, Intl.Segmenter | null>();

function intlWordSegmenter(script: SeaScript): Intl.Segmenter | null {
  const cached = intlWordCache.get(script);
  if (cached !== undefined) return cached;
  let seg: Intl.Segmenter | null = null;
  try {
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
      seg = new Intl.Segmenter(script, { granularity: 'word' });
    }
  } catch {
    seg = null;
  }
  intlWordCache.set(script, seg);
  return seg;
}

function wordSegmenterFor(script: SeaScript): ((span: string) => Iterable<SeaSegment>) | null {
  const intl = intlWordSegmenter(script);
  return intl ? (span) => intl.segment(span) as Iterable<SeaSegment> : null;
}

// ── Break offsets ────────────────────────────────────────────────────────────

/**
 * Enumerate no-space line-break opportunities in `text`, as UTF-16 offsets `i`
 * such that a line break is permitted immediately BEFORE `text[i]`.
 *
 * For DICTIONARY scripts (Thai/Lao/Khmer) the offsets are dictionary WORD starts:
 * boundaries that are (1) interior to a maximal same-class span and (2) the start
 * of a word-like segment — so we never break before intra-script punctuation
 * (Khmer `។`, Thai `๚`). For GRAPHEME-FILL scripts (Myanmar/Tibetan, #961) the
 * offsets are EVERY interior grapheme-cluster boundary (Word's grapheme-fill;
 * no dictionary, no tsheg). Either way the offsets are grapheme-cluster
 * boundaries — a base + stacked/combining cluster is never split — and a
 * SEA↔non-SEA transition is left to the existing logic.
 *
 * Returns `[]` (graceful fallback) when `text` has no no-space character. A
 * dictionary span adds no offsets when `Intl.Segmenter` is unavailable or throws;
 * a grapheme-fill span falls back to code-point boundaries. Segment ONCE per
 * logical string and cache/slice the result; do not call this per line.
 */
export function seaWordBreakOffsets(text: string): number[] {
  if (!containsSeaScript(text)) return [];
  const offsets: number[] = [];
  const len = text.length;
  let i = 0;
  while (i < len) {
    const cp = text.codePointAt(i)!;
    const step = cp > 0xffff ? 2 : 1;
    if (!isSeaScriptCodePoint(cp)) {
      i += step;
      continue;
    }
    // Maximal SAME-CLASS span [i, j): keep dictionary scripts (th/lo/km) and
    // grapheme-fill scripts (my/bo) in separate spans so each gets the right
    // treatment; a dictionary span's dictionary is picked from its first char
    // (ICU still dispatches per-script internally within it).
    const script = seaScriptOf(cp);
    const fill = isGraphemeFillScript(script);
    let j = i + step;
    while (j < len) {
      const c = text.codePointAt(j)!;
      if (!isSeaScriptCodePoint(c)) break;
      if (isGraphemeFillScript(seaScriptOf(c)) !== fill) break;
      j += c > 0xffff ? 2 : 1;
    }
    const span = text.slice(i, j);
    if (fill) {
      // Grapheme-fill (Myanmar/Tibetan): every interior grapheme-cluster boundary
      // is a legal break (Word ground truth, #961). graphemeClusterOffsets is
      // grapheme-safe and degrades to code-point boundaries without a segmenter.
      for (const off of graphemeClusterOffsets(span)) offsets.push(i + off);
    } else {
      const segFn = wordSegmenterFor(script);
      if (segFn) {
        // Accumulate this span's offsets in a temp buffer and commit them only
        // after the iterator completes, so a mid-iteration `segment()` failure
        // leaves NO partial offsets (all-or-nothing graceful fallback per span).
        const spanOffsets: number[] = [];
        try {
          for (const g of segFn(span)) {
            // index>0 keeps the offset interior to the span (never span-start,
            // which is a SEA↔non-SEA edge or the string start). isWordLike keeps
            // it a break-before-a-WORD (never before punctuation). Fake segmenters
            // may omit isWordLike → treat as word-like.
            if (g.index > 0 && (g.isWordLike ?? true)) spanOffsets.push(i + g.index);
          }
          for (const o of spanOffsets) offsets.push(o);
        } catch {
          // segment() failure on this span → add no offsets (graceful fallback).
        }
      }
    }
    i = j;
  }
  return offsets;
}

// ── Script-transition break opportunities (issue #960) ───────────────────────

/**
 * Break-before offsets at the EDGES of every maximal SEA span — the SEA↔non-SEA
 * script transitions that carry NO intervening whitespace (issue #960). Word
 * treats these transitions as line-break opportunities (adjudicated against the
 * Word-exported ground truth sample-45.pdf: a line broke exactly at the Thai→
 * Latin boundary `…ของ | Thailand`, carrying the Latin word whole to the next
 * line; digit groups `1250`/`990` break away from the surrounding Thai the same
 * way). {@link seaWordBreakOffsets} deliberately restricts itself to boundaries
 * INTERIOR to a SEA span, so the transition edges were previously unbreakable —
 * a spaceless mix like `เมืองBangkok…Thailand` had no legal break at the script
 * seam. This adds exactly those seams.
 *
 * Returns UTF-16 offsets `i` (a break is permitted immediately BEFORE `text[i]`),
 * strictly interior (`0 < i < text.length`), ascending and unique. A transition
 * where either side is whitespace is skipped — a space is already an inter-word
 * break, so the seam beside it needs no extra opportunity. Returns `[]` when
 * `text` has no SEA character (byte-identical non-SEA path).
 */
export function seaTransitionOffsets(text: string): number[] {
  if (!containsSeaScript(text)) return [];
  const offsets: number[] = [];
  const len = text.length;
  let prevCp = text.codePointAt(0)!;
  let i = prevCp > 0xffff ? 2 : 1;
  while (i < len) {
    const cp = text.codePointAt(i)!;
    const step = cp > 0xffff ? 2 : 1;
    const prevSea = isSeaScriptCodePoint(prevCp);
    const curSea = isSeaScriptCodePoint(cp);
    // A SEA↔non-SEA change is a transition. Skip when either side is whitespace
    // (the space itself is the break opportunity, handled by the caller).
    if (prevSea !== curSea && !isBreakInertSpaceCp(prevCp) && !isBreakInertSpaceCp(cp)) {
      offsets.push(i);
    }
    prevCp = cp;
    i += step;
  }
  return offsets;
}

/** Code points beside which a SEA↔non-SEA seam must NOT gain a break: real
 *  whitespace (the space itself already IS the break) and the non-breaking space
 *  family (NBSP U+00A0, figure space U+2007, narrow NBSP U+202F, word joiner
 *  U+2060, ZWNBSP/BOM U+FEFF), which are explicitly NON-breaking (UAX#14 GL/WJ). */
function isBreakInertSpaceCp(cp: number): boolean {
  return (
    cp === 0x20 || cp === 0x09 || cp === 0x0a || cp === 0x0d || cp === 0x3000 ||
    cp === 0x00a0 || cp === 0x2007 || cp === 0x202f || cp === 0x2060 || cp === 0xfeff
  );
}

/** Kinsoku sets used to drop line-break positions that would leave a
 *  line-start-forbidden char at a line head or a line-end-forbidden char at a
 *  line tail (ECMA-376 §17.15.1.58–.60). Shape-compatible with the renderers'
 *  `KinsokuRules`; only the two membership sets are consulted. */
export interface SeaMixedKinsoku {
  /** When `false`, the document turned kinsoku OFF (§17.3.1.16 `<w:kinsoku
   *  w:val="0"/>`) and NO position is dropped — matching the CJK path's own
   *  `if (!rules.enabled) return` short-circuit. Absent/`true` ⇒ filter. */
  enabled?: boolean;
  lineStartForbidden: ReadonlySet<number>;
  lineEndForbidden: ReadonlySet<number>;
}

/**
 * The UNIFIED break-before offset set for a single run/segment that contains SEA
 * script but may also mix Latin, digits and CJK (issue #960). Unions three
 * opportunity sources so each script keeps its own rule inside one contiguous
 * token:
 *   • SEA dictionary word boundaries — {@link seaWordBreakOffsets};
 *   • SEA↔non-SEA transitions — {@link seaTransitionOffsets};
 *   • CJK per-character boundaries when `opts.cjk` — every position whose char,
 *     or the char before it, is a {@link isCjkBreakChar} ideograph, so the CJK
 *     side breaks at each character exactly as its dedicated path does.
 *
 * When `opts.kinsoku` is supplied, positions that would orphan a
 * line-start-forbidden char at a line head, or leave a line-end-forbidden char
 * at a line tail, are removed (the offset-set equivalent of the CJK path's
 * retract). This is what lets a mixed CJK+SEA token — which previously fell
 * wholly onto the CJK path and so treated its SEA interior as unbreakable — wrap
 * with BOTH the CJK per-char opportunities and the SEA dictionary/transition
 * opportunities merged.
 *
 * Offsets are ascending, unique, strictly interior. For a pure-SEA token with no
 * kinsoku char this equals {@link seaWordBreakOffsets} (byte-identical wrap).
 */
export function seaMixedBreakOffsets(
  text: string,
  opts?: { cjk?: boolean; kinsoku?: SeaMixedKinsoku },
): number[] {
  if (!containsSeaScript(text)) return [];
  const len = text.length;
  // Dictionary boundaries are ALWAYS grapheme-cluster boundaries (ICU never
  // splits a base + mark), so seed the set with them directly.
  const set = new Set<number>(seaWordBreakOffsets(text));
  // The transition and CJK offsets are enumerated per code point, so they could
  // land INSIDE a grapheme cluster (a base + variation selector / ZWJ / tone
  // mark). Collect them separately and keep only those that coincide with a
  // grapheme boundary — a code-point split there would tear the cluster.
  const graphemeSafe = new Set<number>(graphemeClusterOffsets(text));
  const addGraphemeSafe = (o: number): void => {
    if (graphemeSafe.has(o)) set.add(o);
  };
  for (const o of seaTransitionOffsets(text)) addGraphemeSafe(o);
  if (opts?.cjk) {
    let prevCp = text.codePointAt(0)!;
    let i = prevCp > 0xffff ? 2 : 1;
    while (i < len) {
      const cp = text.codePointAt(i)!;
      // A break may fall before an ideograph or after one (CJK is break-eligible
      // on both edges; the grapheme test above drops a variation-selector tear,
      // and the kinsoku pass below removes the illegal ones).
      if (isCjkBreakChar(cp) || isCjkBreakChar(prevCp)) addGraphemeSafe(i);
      prevCp = cp;
      i += cp > 0xffff ? 2 : 1;
    }
  }
  const k = opts?.kinsoku;
  // §17.3.1.16 — a document that turned kinsoku OFF drops NO position (mirrors
  // the CJK split path's `if (!rules.enabled) return`).
  const filterKinsoku = k != null && k.enabled !== false;
  const out: number[] = [];
  for (const o of set) {
    if (o <= 0 || o >= len) continue;
    if (filterKinsoku) {
      const at = text.codePointAt(o)!;
      const before = codePointEndingAt(text, o);
      if (before !== undefined && k!.lineEndForbidden.has(before)) continue;
      if (k!.lineStartForbidden.has(at)) continue;
    }
    out.push(o);
  }
  out.sort((a, b) => a - b);
  return out;
}

/** The code point that ENDS immediately before UTF-16 offset `o` (handles a
 *  trailing surrogate pair). `undefined` when `o` is at or before the start. */
function codePointEndingAt(text: string, o: number): number | undefined {
  if (o <= 0) return undefined;
  const lo = text.charCodeAt(o - 1);
  if (lo >= 0xdc00 && lo <= 0xdfff && o >= 2) {
    const hi = text.charCodeAt(o - 2);
    if (hi >= 0xd800 && hi <= 0xdbff) return text.codePointAt(o - 2)!;
  }
  return lo;
}

// ── Greedy whole-prefix line fit (shared kernel) ─────────────────────────────

/**
 * Greedy line fit for no-space breaking (dictionary words or grapheme clusters),
 * shared by all three renderers so the break decision is defined once. `offsets`
 * are the legal break-before positions from {@link seaWordBreakOffsets}.
 *
 * Returns the largest `end` with `start < end <= text.length` such that
 *   (a) `end === text.length` OR `end` is a member of `offsets` (a legal word
 *       boundary), AND
 *   (b) `measure(text.slice(start, end)) <= avail`.
 * If not even the first candidate word fits within `avail`, returns `start`
 * (no progress) so the caller can wrap to a fresh line first, or — on an already
 * empty line — fall back to a grapheme-safe emergency split.
 *
 * By default EVERY candidate is measured (no early return): advance width is NOT
 * guaranteed monotone in the prefix length because Word/PowerPoint allow NEGATIVE
 * character spacing (`w:spacing` / `a:spc`), so a longer prefix can fit after a
 * shorter one overflowed. We keep the largest boundary (or the full remainder)
 * that fits. This full scan is O(#offsets) measures per call — fine for the sparse
 * dictionary-word offsets of Thai/Lao/Khmer.
 *
 * `assumeMonotone` switches to an O(log #offsets) BINARY SEARCH for callers whose
 * advance IS monotone in the prefix length — the grapheme-fill scripts
 * Myanmar/Tibetan (#961), whose `offsets` are EVERY grapheme boundary (dense: one
 * per cluster), where a per-line full scan would be O(n²) down a long run. This
 * matches the monotone assumption {@link fitCJKPrefix} already makes for CJK. Do
 * NOT set it for a run that can carry glyph-overlapping negative spacing (the
 * dictionary path keeps the safe full scan).
 *
 * `offsets` MUST be sorted ascending (as produced by {@link seaWordBreakOffsets}
 * or {@link graphemeClusterOffsets}). `measure` supplies the caller's exact
 * advance model (so measure==paint). Pure — no canvas/DOM.
 */
export function fitSeaWordPrefix(
  text: string,
  offsets: readonly number[],
  start: number,
  avail: number,
  measure: (sub: string) => number,
  assumeMonotone = false,
): number {
  const len = text.length;
  if (start >= len) return start;
  if (assumeMonotone) {
    // Monotone advance: "fits" is downward-closed over the boundary position, so
    // the whole remainder is the max if it fits, else binary-search the interior
    // offsets (ascending) for the largest fitting one. Offsets ≤ start form a
    // prefix and ≥ len a suffix of the sorted array; the search skips both.
    if (measure(text.slice(start, len)) <= avail) return len;
    let lo = 0;
    let hi = offsets.length - 1;
    let ans = start;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const b = offsets[mid];
      if (b <= start) lo = mid + 1;
      else if (b >= len) hi = mid - 1;
      else if (measure(text.slice(start, b)) <= avail) {
        ans = b;
        lo = mid + 1;
      } else hi = mid - 1;
    }
    return ans;
  }
  let best = start;
  for (const b of offsets) {
    if (b <= start || b >= len) continue;
    if (measure(text.slice(start, b)) <= avail) best = b;
  }
  if (measure(text.slice(start, len)) <= avail) best = len;
  return best;
}

// ── Grapheme clusters (emergency over-long-word split) ───────────────────────

let graphemeSegmenter: Intl.Segmenter | null | undefined;
// Characters whose Grapheme_Cluster_Break value is Control, CR, LF or Other
// (UAX #29): Basic Latin and Latin-1 Supplement, CJK Symbols and Punctuation
// without the Extend marks U+302A..U+302F, Hiragana and Katakana without the
// Extend sound marks U+3099..U+309A, CJK Unified Ideographs and Extension A,
// and the Halfwidth and Fullwidth Forms up to U+FF9D (U+FF9E..U+FF9F are
// Extend). None is Extend, ZWJ, SpacingMark, Prepend, Hangul jamo or
// Regional_Indicator, so GB3 (CR × LF) is the only rule that joins two of
// them.
const SINGLE_UNIT_GRAPHEME_TEXT =
  /^[\u0000-\u00ff\u3000-\u3029\u3030-\u303f\u3041-\u3096\u309b-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uff01-\uff9d]*$/;

/** The realm's shared extended-grapheme segmenter, created on first use. */
export function getGraphemeSegmenter(): Intl.Segmenter | null {
  if (graphemeSegmenter !== undefined) return graphemeSegmenter;
  let seg: Intl.Segmenter | null = null;
  try {
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
      seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    }
  } catch {
    seg = null;
  }
  graphemeSegmenter = seg;
  return seg;
}

/**
 * Grapheme-cluster break-before offsets (UTF-16, interior — excludes 0 and
 * `text.length`). Used for the last-resort emergency split of a single SEA word
 * wider than the whole line/cell: a code-point split would tear a base +
 * combining/tone mark (both are BMP, so "SEA is BMP therefore safe" is false),
 * whereas a grapheme boundary keeps the cluster intact. Falls back to code-point
 * boundaries when `Intl.Segmenter` is unavailable OR `segment()`/its iterator
 * throws (a `small-icu` build, an exotic runtime) — never propagates, so the
 * grapheme-fill wrap (#961) degrades to a safe code-point split instead of failing.
 */
export function graphemeClusterOffsets(text: string): number[] {
  // UAX #29 fast path: text made only of the characters above, without a
  // CR LF pair, has a boundary between every pair of code units. The segmenter
  // and the code-point fallback below both return exactly that. Skipping them
  // avoids constructing an ICU break iterator (several ms on the first call in
  // each realm) and one segment object per character.
  if (SINGLE_UNIT_GRAPHEME_TEXT.test(text) && !text.includes('\r\n')) {
    const out = new Array<number>(Math.max(0, text.length - 1));
    for (let i = 1; i < text.length; i += 1) out[i - 1] = i;
    return out;
  }
  const seg = getGraphemeSegmenter();
  if (seg) {
    try {
      const out: number[] = [];
      for (const g of seg.segment(text)) {
        if (g.index > 0) out.push(g.index);
      }
      return out;
    } catch {
      // fall through to the code-point boundaries below (graceful fallback)
    }
  }
  const out: number[] = [];
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    i += cp > 0xffff ? 2 : 1;
    if (i < text.length) out.push(i);
  }
  return out;
}
