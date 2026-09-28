import { isCjkBreakChar, isLatinWordCodePoint } from '../text/cjk-ranges.js';
import { isUax14NoBreakPair } from '../text/line-break.js';
import { kinsokuAdjustedSplit, DEFAULT_KINSOKU_RULES } from '../text/kinsoku/index.js';
import { containsSeaScript, isGraphemeFillText, seaMixedBreakOffsets,
  fitSeaWordPrefix, graphemeClusterOffsets } from '../text/sea-break.js';

/** SpreadsheetML cell host hook: `xf@wrapText` and rich shared-string runs are
 * not `a:txBody`. Their token/overflow policy differs from shape paragraphs,
 * while using the same core Unicode and kinsoku primitives. Font resolution
 * remains in the XLSX adapter. */
/**
 * Apply Japanese line-breaking (kinsoku, 禁則処理) at a wrap boundary.
 *
 * When the wrapper decides to break, it has the code points already committed
 * to the line being closed (`lineCps`) and the code points that will lead the
 * next line (`nextCps`, the overflowing token). Excel — like Word and
 * PowerPoint — forbids a wrapped line from STARTING with a 行頭禁則 char
 * (、。」）…) or ENDING with a 行末禁則 char (「（…), per ECMA-376
 * §17.15.1.58–.60. We delegate the retraction to the shared core engine
 * `kinsokuAdjustedSplit`, which pulls the offending boundary's preceding code
 * point(s) down onto the next line (追い出し).
 *
 * Returns the number of trailing code points of `lineCps` that must move down
 * to lead the next line (ahead of `nextCps`). `0` means the greedy break was
 * already legal — so plain CJK with no forbidden chars at the boundary is
 * unchanged (no regression). `minSplit = 1` keeps ≥1 code point on the closed
 * line; an all-forbidden run falls back to no retraction (never empties a line,
 * never hangs).
 */
function kinsokuRetractCount(lineCps: string[], nextCps: string[]): number {
  if (lineCps.length === 0 || nextCps.length === 0) return 0;
  const combined = [...lineCps, ...nextCps];
  const splitAt = lineCps.length;
  const adj = kinsokuAdjustedSplit(combined, splitAt, DEFAULT_KINSOKU_RULES, 1);
  return splitAt - adj;
}

/**
 * Extend a per-code-point kinsoku retract so it never TEARS a Latin word.
 *
 * `kinsokuRetractCount` retracts glyph-by-glyph — correct for CJK, where every
 * character is a break opportunity, but wrong for Latin: a non-starter (comma,
 * period, …, UAX#14 LB13) overflowing after "system" retracts a single "m" →
 * "syste" / "m,". Latin has no mid-word break opportunity, so when the retract
 * boundary sits between two {@link isLatinWordCodePoint} characters, pull it back
 * to the last whitespace (the real break) so the WHOLE word moves down ahead of
 * the non-starter. If the line is one unbroken word (no whitespace to retract
 * to), keep the original retract rather than empty the line — an over-long
 * single word is handled by the normal overflow path. NOTE: the retract is still
 * capped at the last segment by the caller, so a word split across runs by a
 * formatting change splits at that seam (the comma stays glued to the tail).
 */
function extendLatinWordRetract(lineCps: string[], retract: number): number {
  let r = retract;
  while (r < lineCps.length) {
    const keep = lineCps[lineCps.length - r - 1]; // last char staying on the line
    const move = lineCps[lineCps.length - r];     // first char moving down
    const keepCp = keep?.codePointAt(0);
    const moveCp = move?.codePointAt(0);
    if (keepCp !== undefined && moveCp !== undefined
        && isLatinWordCodePoint(keepCp) && isLatinWordCodePoint(moveCp)) r++;
    else break;
  }
  return r >= lineCps.length ? retract : r;
}

/**
 * Return the supported UAX #14 no-break suffix of the current line that must
 * move before `nextCps`. The predicate is deliberately one-way: walking stops
 * on false even when a deferred rule might also prohibit that earlier boundary.
 * Returning zero for a whole-line sequence preserves the existing emergency
 * overflow behavior and avoids emitting an empty soft-wrapped line.
 */
function uaxNoBreakRetractCount(lineCps: string[], nextCps: string[]): number {
  if (lineCps.length === 0 || nextCps.length === 0) return 0;
  const nextCp = nextCps[0].codePointAt(0);
  let firstMoved = lineCps.length - 1;
  const lastCp = lineCps[firstMoved].codePointAt(0);
  if (
    lastCp === undefined ||
    nextCp === undefined ||
    lastCp === 0x200b ||
    nextCp === 0x200b ||
    !isUax14NoBreakPair(lastCp, nextCp)
  ) return 0;

  while (firstMoved > 0) {
    const prevCp = lineCps[firstMoved - 1].codePointAt(0);
    const movedCp = lineCps[firstMoved].codePointAt(0);
    if (
      prevCp === undefined ||
      movedCp === undefined ||
      !isUax14NoBreakPair(prevCp, movedCp)
    ) break;
    firstMoved--;
  }

  return firstMoved === 0 ? 0 : lineCps.length - firstMoved;
}

/** Word-wrap a single paragraph (no embedded \n). Unlike a naive
 *  `split(' ')`, CJK characters are treated as individual break opportunities
 *  so that Japanese headings like "夏休みアクティビティ カレンダー 2026"
 *  actually wrap inside a merged cell. ECMA-376 doesn't spec the break
 *  algorithm but this matches what Excel renders on the same input.
 *
 *  At each break we additionally apply kinsoku (`kinsokuRetractCount`) so a
 *  wrapped line never starts with 、。」 or ends with 「（ (ECMA-376
 *  §17.15.1.58–.60), matching Excel's East-Asian wrapping. */
export function wrapSpreadsheetCellParagraph(
  paragraph: string, maxWidth: number,
  measureWidth: (value: string) => number,
): string[] {
  const lines: string[] = [];
  // Tokenise: runs of non-space non-CJK, single ASCII-space runs, individual
  // CJK characters. Then greedy-fit each token onto the current line.
  const tokens: string[] = [];
  let i = 0;
  while (i < paragraph.length) {
    const ch = paragraph[i];
    const cp = ch.codePointAt(0) ?? 0;
    if (isCjkBreakChar(cp)) {
      tokens.push(ch);
      i += cp > 0xFFFF ? 2 : 1;
    } else if (ch === ' ') {
      let j = i;
      while (j < paragraph.length && paragraph[j] === ' ') j++;
      tokens.push(paragraph.slice(i, j));
      i = j;
    } else {
      let j = i;
      while (j < paragraph.length) {
        const c = paragraph[j];
        const p = c.codePointAt(0) ?? 0;
        if (c === ' ' || isCjkBreakChar(p)) break;
        j += p > 0xFFFF ? 2 : 1;
      }
      const word = paragraph.slice(i, j);
      // SEA (Thai/Lao/Khmer) dictionary breaking (issue #797): these scripts have
      // no inter-word spaces, so this whole word-run is one token. Split it at
      // segmenter word boundaries into sub-word tokens so the greedy fitter below
      // wraps it at legal points. Wrapped lines re-concatenate into one drawn
      // string, so this only ADDS break opportunities (measure==paint). Non-SEA
      // words and SEA words with no usable break stay a single token.
      // Issue #797 / #960 — dictionary word boundaries UNIONED with the no-space
      // SEA↔non-SEA script transitions (Thai↔Latin/digit), so a price like
      // "…1250…" or an embedded Latin word can wrap away from the surrounding
      // Thai. CJK is already its own token here (split above), so the mixed CJK
      // path is not needed.
      const seaBreaks = containsSeaScript(word) ? seaMixedBreakOffsets(word) : null;
      if (seaBreaks && seaBreaks.length > 0) {
        let s = 0;
        for (const b of seaBreaks) { tokens.push(word.slice(s, b)); s = b; }
        tokens.push(word.slice(s));
      } else {
        tokens.push(word);
      }
      i = j;
    }
  }
  let current = '';
  for (const tok of tokens) {
    if (current === '') { current = tok; continue; }
    const candidate = current + tok;
    if (measureWidth(candidate) <= maxWidth) {
      current = candidate;
    } else {
      // Token doesn't fit at the end of the current line — break here.
      // Leading spaces at the start of the next line are dropped (matches
      // Excel: wrapped-continuation lines don't preserve the space that
      // caused the break).
      let nextLead = tok.replace(/^ +/, '');
      if (nextLead === '') nextLead = tok; // all-space token (preserve width on its own line)
      // Apply kinsoku at the boundary: retract trailing code points of the
      // line being closed so it does not end with a 行末禁則 char and the
      // next line does not start with a 行頭禁則 char.
      const lineCps = [...current];
      const retract = kinsokuRetractCount(lineCps, [...nextLead]);
      if (retract > 0) {
        const keep = lineCps.length - retract;
        lines.push(lineCps.slice(0, keep).join(''));
        current = lineCps.slice(keep).join('') + nextLead;
      } else {
        lines.push(current);
        current = nextLead;
      }
    }
  }
  lines.push(current);
  return lines;
}

export interface SpreadsheetCellRichSeg<F extends { size: number }> {
  text: string;
  font: F;
  width: number; // px
}

export interface SpreadsheetCellRichLine<F extends { size: number }> {
  segments: SpreadsheetCellRichSeg<F>[];
  maxFontSize: number; // pt (line-height source)
  /** 0-based index of the LF-delimited paragraph (hard-break region) this line
   *  belongs to. Soft-wrapped continuation lines share their paragraph's index;
   *  it advances only at a hard break. Indexes the per-paragraph bidi base
   *  direction the wrap draw path resolves — UAX#9: a soft wrap does NOT start a
   *  new paragraph, but a hard break does. */
  para: number;
}

/**
 * Layout rich text runs into wrapped lines. Each run is split into words (and
 * CJK characters for granular wrapping). Per-run font is preserved so measurement
 * and drawing use the correct font.
 *
 * Runs are inline and share the cell width (ECMA-376 §18.4.4 r / §18.4.8 si /
 * §18.4.9 sst); wrapText (§18.8.1) breaks at word boundaries (ASCII spaces) and
 * at any CJK code point boundary, and a hard break (LF) starts a new line.
 *
 * An empty value returns `[]` (no fabricated line). This deliberately differs
 * from the plain-text `wrapTextLines`, whose `split('\n')` yields `['']`: an
 * empty cell has no glyphs, so reserving a line for it would only mis-anchor the
 * (non-existent) text.
 */
export function layoutSpreadsheetCellRichLines<R extends { text: string }, F extends { size: number }>(
  ctx: CanvasRenderingContext2D,
  runs: R[],
  baseFont: F,
  maxWidth: number,
  applyRunFont: (base: F, run: R) => F,
  setDrawFont: (font: F, text: string) => void,
  defaultFontSize: number,
): SpreadsheetCellRichLine<F>[] {
  const lines: SpreadsheetCellRichLine<F>[] = [];
  let cur: SpreadsheetCellRichSeg<F>[] = [];
  let curW = 0;
  let curMaxSize = 0;
  // Size (pt) of the nearest preceding text run — the height source for a blank
  // line, which has no segment of its own. Mirrors drawShapeText's `lastTextPt`
  // seed (PR #583); falls back to the cell's base font.
  let lastTextPt = baseFont.size;
  // 0-based index of the current LF-delimited paragraph. Advances only at a hard
  // break (not at a soft wrap), so every line records which paragraph it belongs
  // to — the wrap draw path resolves a Context base direction per paragraph.
  let paraIdx = 0;

  // `flush` drops an empty region — used at soft-wrap (kinsoku) breaks, where a
  // line carried wholly to the next line must not leave a blank behind.
  const flush = () => {
    if (cur.length === 0) return;
    lines.push({ segments: cur, maxFontSize: curMaxSize, para: paraIdx });
    cur = []; curW = 0; curMaxSize = 0;
  };

  // `flushRegion` emits an empty region as a blank line — used at a hard break
  // (LF) or end-of-value. ECMA-376 §18.8.1 (wrapText): each line of a multi-line
  // cell, including a blank one from consecutive / leading / trailing breaks,
  // reserves one single-line height (the cell analog of PR #583 / docx #582).
  const flushRegion = () => {
    if (cur.length === 0) {
      lines.push({ segments: [], maxFontSize: lastTextPt || defaultFontSize, para: paraIdx });
      return;
    }
    flush();
  };

  const push = (text: string, font: F) => {
    if (!text) return;
    lastTextPt = font.size; // nearest preceding text size, for the next blank line
    // Measure at the *draw* font so a super/subscript token reserves its reduced
    // (~65%) glyph width; the segment keeps the run's full size for line height.
    setDrawFont(font, text);
    const w = ctx.measureText(text).width;
    if (cur.length > 0 && curW + w > maxWidth) {
      // Kinsoku at the wrap boundary (ECMA-376 §17.15.1.58–.60): retract
      // trailing code points of the line being closed so it does not end with
      // a 行末禁則 char and the next line (led by `text`) does not start with a
      // 行頭禁則 char. The retracted code points live at the end of the last
      // segment — split that segment (keeping its font), re-measure both parts
      // with the segment's font, and carry the trailing part down to lead the
      // next line.
      const lineCps = cur.flatMap((s) => [...s.text]);
      let retract = kinsokuRetractCount(lineCps, [...text]);
      // UAX#14 LB13: a per-glyph retract would tear a Latin word (e.g. a comma in
      // a separate run overflowing after "system" → "syste" / "m,"). Pull the
      // retract back to the last whitespace so the whole word rides down with the
      // non-starter; CJK boundaries (move char is CJK) are left untouched.
      if (retract > 0) {
        retract = extendLatinWordRetract(lineCps, retract);
      } else if (
        // SEA (Thai/Lao/Khmer) dictionary tailoring wins over the LB1 SA→AL
        // default on BOTH sides: guard the incoming `text` AND the last segment
        // that would be retracted, so the UAX #14 pair predicate never
        // suppresses a SEA word boundary (mirror the DOCX buildSegments
        // prev/cur guard). Retraction is capped to the last segment below, so
        // checking it is the precise preceding-side test.
        !containsSeaScript(text) &&
        !containsSeaScript(cur[cur.length - 1]?.text ?? '') &&
        !/^\s/u.test(text) &&
        !/\s$/u.test(lineCps.at(-1) ?? '')
      ) {
        retract = uaxNoBreakRetractCount(lineCps, [...text]);
      }
      const last = cur[cur.length - 1];
      const lastCps = [...last.text];
      // Only retract within the last segment to preserve each run's font; the
      // single-run CJK case (one segment per line region) is fully covered.
      if (retract > lastCps.length) retract = lastCps.length;
      let carry: SpreadsheetCellRichSeg<F> | null = null;
      if (retract > 0) {
        const keepCps = lastCps.slice(0, lastCps.length - retract);
        const moveCps = lastCps.slice(lastCps.length - retract);
        if (keepCps.length === 0) {
          // The whole last segment moves down — drop it from the closing line.
          cur.pop();
        } else {
          const keepText = keepCps.join('');
          setDrawFont(last.font, keepText);
          last.text = keepText;
          last.width = ctx.measureText(keepText).width;
        }
        const moveText = moveCps.join('');
        setDrawFont(last.font, moveText);
        carry = { text: moveText, font: last.font, width: ctx.measureText(moveText).width };
      }
      flush();
      if (carry) {
        cur.push(carry);
        curW += carry.width;
        if (carry.font.size > curMaxSize) curMaxSize = carry.font.size;
      }
      setDrawFont(font, text); // restore for the incoming token below
    }
    cur.push({ text, font, width: w });
    curW += w;
    if (font.size > curMaxSize) curMaxSize = font.size;
  };

  // Issue #797 — push a SEA (Thai/Lao/Khmer) token, breaking it at segmenter word
  // boundaries. Unlike the plain `wrapParagraphLines`, the rich draw path paints
  // every segment separately, so each fitted line-piece is pushed as ONE
  // contiguous string (via `push`) to keep measure==paint. A single word wider
  // than the cell falls back to a grapheme-safe emergency split.
  const pushSeaToken = (text: string, font: F): void => {
    // #797 dictionary boundaries ∪ #960 SEA↔non-SEA transitions (CJK is a
    // separate token in this path, so no mixed-CJK offsets are needed here).
    const seaBreaks = seaMixedBreakOffsets(text);
    if (seaBreaks.length === 0) { push(text, font); return; }
    setDrawFont(font, text);
    const measureSub = (sub: string): number => ctx.measureText(sub).width;
    // Grapheme-fill runs (Myanmar/Tibetan, #961) have dense per-cluster offsets:
    // O(log n) monotone binary-search fit. Dictionary runs keep the full scan.
    const monotone = isGraphemeFillText(text);
    const N = text.length;
    let start = 0;
    while (start < N) {
      const avail = maxWidth - curW;
      let end = fitSeaWordPrefix(text, seaBreaks, start, avail, measureSub, monotone);
      if (end <= start) {
        if (curW > 0) { flush(); continue; } // wrap first, retry on an empty line
        const firstWordEnd = seaBreaks.find((b) => b > start) ?? N;
        const firstWord = text.slice(start, firstWordEnd);
        const graphemes = graphemeClusterOffsets(firstWord);
        let g = fitSeaWordPrefix(firstWord, graphemes, 0, avail, measureSub, monotone);
        if (g <= 0) g = graphemes.length > 0 ? graphemes[0] : firstWord.length;
        end = start + g;
      }
      push(text.slice(start, end), font); // the piece fits → append (no re-split)
      start = end;
      if (start < N) flush();
    }
  };

  for (const run of runs) {
    const font = applyRunFont(baseFont, run);
    // Tokenize: runs of non-space latin, spaces, or individual CJK chars
    const tokens: string[] = [];
    let i = 0;
    while (i < run.text.length) {
      const ch = run.text[i];
      const cp = ch.codePointAt(0) ?? 0;
      if (cp === 0x000A) {
        // Explicit newline: force break
        tokens.push('\n'); i += 1;
      } else if (isCjkBreakChar(cp)) {
        tokens.push(ch);
        i += cp > 0xFFFF ? 2 : 1;
      } else if (ch === ' ') {
        let j = i;
        while (j < run.text.length && run.text[j] === ' ') j++;
        tokens.push(run.text.slice(i, j));
        i = j;
      } else {
        let j = i;
        while (j < run.text.length) {
          const c = run.text[j];
          const p = c.codePointAt(0) ?? 0;
          if (c === ' ' || c === '\n' || isCjkBreakChar(p)) break;
          j += p > 0xFFFF ? 2 : 1;
        }
        tokens.push(run.text.slice(i, j));
        i = j;
      }
    }
    for (const tok of tokens) {
      // A hard break closes the current paragraph region and opens the next, so
      // the following lines record the new paragraph index. A soft wrap (handled
      // inside `push`) keeps the same index — UAX#9 P1: only a hard break starts
      // a new bidi paragraph.
      if (tok === '\n') { flushRegion(); paraIdx++; }
      else if (containsSeaScript(tok)) pushSeaToken(tok, font);
      else push(tok, font);
    }
  }
  // Trailing region. If the value ended with a break, `cur` is empty but a line
  // was already produced, so a trailing blank line is reserved; a value with no
  // content and no breaks (no segments, no prior line) produces nothing.
  if (cur.length > 0 || lines.length > 0) flushRegion();
  return lines;
}
