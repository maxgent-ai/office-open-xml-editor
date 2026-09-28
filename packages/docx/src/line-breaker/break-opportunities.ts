import { justifiedCandidateFitWidth } from './justify-fit.js';
import {
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  isGraphemeFillText,
  isDictionarySeaText,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
} from '@silurus/ooxml-core';
import { EAST_ASIAN_RE, nextTabStop } from '../layout/text.js';
import {
  wordIsOverflowPunctuation,
  wordIdeographicSpaceLineEndAllowanceCount,
} from '../layout/line-compatibility.js';
import {
  type LayoutImageSeg,
  type LayoutMathSeg,
  type LayoutSeg,
  type LayoutTabSeg,
  type LayoutTextSeg,
} from './model.js';
import {
  RESET_SLICED_TEXT_MEASUREMENT,
  charScaleFactor,
  charSpacingDeltaPx,
  hardJoinPrefixEnd,
  legalTextSplitAtOrBefore,
  segAdvanceWidth,
  segmentCharacterGridDeltaPx,
  slicedTextMetadata,
  snapToCharsClass,
} from './advance.js';
import { tabAlignmentRole } from './tabs.js';
import { buildFont } from './font-routes.js';
import {
  extendThroughTrailingIdeographicSpaces,
  hasCJKBreakOpportunity,
  rebaseSeaBreaks,
} from './text-runs.js';
import { fitCJKPrefix, hasEastAsianVisiblePredecessor } from './fit-search.js';
import type { PassOperationState } from './pass-operations.js';

/** The iterator reads only its declared slice of the explicit pass state. */
export type BreakOpportunityIteratorContext = Pick<
  PassOperationState,
  | 'breakerState'
  | 'flush'
  | 'baseRtl'
  | 'addToLine'
  | 'scale'
  | 'firstIndent'
  | 'tabOriginPx'
  | 'maxWidth'
  | 'marginRightPx'
  | 'tabFollowWidth'
  | 'measureText'
  | 'verticalInkExtra'
  | 'characterGrid'
  | 'tabStops'
  | 'defaultTabPt'
  | 'tabFollowingMetrics'
  | 'availW'
  | 'setMeasureFont'
  | 'fontFamilyClasses'
  | 'measurement'
  | 'textSegmentBox'
  | 'prospectiveSnapAdvance'
  | 'segAdvance'
  | 'strAdvance'
  | 'isJustified'
  | 'stretchLastLine'
  | 'overflowPunct'
  | 'sameLatinSpaceFace'
  | 'fitsMeasuredWidth'
  | 'fitHomogeneousLatinSpaces'
  | 'appendQueuedIdeographicSpaceSegment'
  | 'emergencyTextSplit'
  | 'effectiveFontPx'
  | 'ctx'
  | 'verticalGlyphMeasurement'
  | 'kinsoku'
  | 'strNaturalAdvance'
  | 'retractCurrentLineForLeadingKinsoku'
  | 'keepLeadingKinsokuWithCurrentLine'
  | 'externalLinkSyntaxSplit'
  | 'queueEmergencyTail'
>;

/** Consume one prepared queue in source order, applying all legal break paths. */
export function iterateBreakOpportunities(context: BreakOpportunityIteratorContext): void {
  const { breakerState, flush } = context;
  while (breakerState.queue.length > 0) {
    const seg = breakerState.queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      // The line being flushed ends at a MANUAL break (§17.3.3.1) — mark it so a
      // justified paragraph left-aligns it like its final line (§17.18.44).
      flush(seg.fontSize, true);
      breakerState.trailingBreakFontSize = seg.fontSize;
      continue;
    }
    breakerState.trailingBreakFontSize = null;

    // ── Tab segment ──────────────────────────────────────
    if ('isTab' in seg) {
      processTabSegment(context, seg);
      continue;
    }

    // ── Image segment ────────────────────────────────────
    if ('imagePath' in seg) {
      processImageSegment(context, seg);
      continue;
    }

    // ── Math segment ─────────────────────────────────────
    if ('math' in seg) {
      processMathSegment(context, seg);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    processTextSegment(context, seg as LayoutTextSeg);
  }
}

function processTextSegment(context: BreakOpportunityIteratorContext, seg: LayoutTextSeg): void {
  const {
    breakerState,
    flush,
    addToLine,
    maxWidth,
    characterGrid,
    availW,
    textSegmentBox,
    prospectiveSnapAdvance,
    segAdvance,
    strAdvance,
    isJustified,
    stretchLastLine,
    overflowPunct,
    sameLatinSpaceFace,
  } = context;
  const s = seg as LayoutTextSeg;
  const segmentBox = textSegmentBox(s);
  const w = segmentBox.width;
  const prospectiveWidth = prospectiveSnapAdvance(s, w);
  const h = segmentBox.height;
  const asc = segmentBox.ascent;
  const desc = segmentBox.descent;
  const paragraphFinalIdeographicSpaceTail = s.paragraphFinalIdeographicSpaceTail === true;
  const paragraphFinalIdeographicSpaceCount = s.paragraphFinalIdeographicSpaceCount ?? 0;
  const paragraphFinalIdeographicSpaceLocalCount = s.paragraphFinalIdeographicSpaceLocalCount ?? 0;
  const visibleBeforeParagraphFinalTail = paragraphFinalIdeographicSpaceTail
    ? s.text.slice(0, Math.max(0, s.text.length - paragraphFinalIdeographicSpaceLocalCount))
    : s.text;
  if (
    paragraphFinalIdeographicSpaceTail &&
    paragraphFinalIdeographicSpaceCount > 1 &&
    visibleBeforeParagraphFinalTail.length > 0
  ) {
    const visibleSegment: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: visibleBeforeParagraphFinalTail,
      paragraphFinalIdeographicSpaceTail: undefined,
      paragraphFinalIdeographicSpaceLocalCount: undefined,
      paragraphFinalIdeographicSpaceCount: undefined,
      paragraphFinalIdeographicSpaceTailStart: undefined,
      measuredWidth: 0,
      ...slicedTextMetadata(s, 0, visibleBeforeParagraphFinalTail.length),
    };
    const trailingSegment: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: s.text.slice(visibleBeforeParagraphFinalTail.length),
      paragraphFinalIdeographicSpaceLocalCount,
      joinPrev: undefined,
      hardJoinPrev: undefined,
      paragraphFinalIdeographicSpaceTailStart: true,
      measuredWidth: 0,
      ...slicedTextMetadata(s, visibleBeforeParagraphFinalTail.length, s.text.length),
      src: s.src
        ? {
            segIndex: s.src.segIndex,
            charOffset: s.src.charOffset + visibleBeforeParagraphFinalTail.length,
          }
        : undefined,
    };
    breakerState.queue.unshift(trailingSegment);
    breakerState.queue.unshift(visibleSegment);
    return;
  }
  if (
    paragraphFinalIdeographicSpaceTail &&
    /^\u3000+$/u.test(s.text) &&
    s.paragraphFinalIdeographicSpaceTailStart === true
  ) {
    const currentLineHasVisibleText = breakerState.currentLine.some(
      (candidate) => 'text' in candidate && /[^\u3000]/u.test(candidate.text),
    );
    if (currentLineHasVisibleText) {
      let trailingTailWidth = w;
      for (const candidate of breakerState.queue) {
        if (!('text' in candidate) || candidate.paragraphFinalIdeographicSpaceTail !== true) break;
        trailingTailWidth += segAdvance(candidate);
      }
      if (breakerState.currentWidth + trailingTailWidth > availW()) {
        flush(undefined, false, s.src);
        breakerState.queue.unshift(s);
        return;
      }
    }
  }

  // ECMA-376 §17.3.2.14: a fit region is an atomic fixed-width cell. The
  // first segment judges the WHOLE resolved region; after an optional flush,
  // every member is added without entering the CJK/overlong-word split paths.
  // This also handles a target wider than the line: it overflows as one unit
  // instead of violating the required internal non-wrap boundary.
  if (s.fitTextRegionIndex !== undefined) {
    if (s.fitTextRegionStart) {
      let regionWidth = w;
      for (const queued of breakerState.queue) {
        if (!('text' in queued) || queued.fitTextRegionIndex !== s.fitTextRegionIndex) break;
        regionWidth += segAdvance(queued);
      }
      if (
        breakerState.currentLine.length > 0 &&
        breakerState.currentWidth + regionWidth > availW()
      ) {
        flush(undefined, false, s.src);
      }
    }
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    return;
  }
  // A terminal separator may collapse when this word becomes line-final;
  // visible glyphs still need to fit at their natural measured advance.
  const trimmed = s.text.replace(/ +$/, '');
  // Subtract the full-model advance of the trimmed text (not the natural width)
  // so the grid delta, w:w scale and w:spacing pitch on the retained glyphs all
  // cancel and trailingSpaceW is the bare trailing-space advance — keeping `w`
  // and `wForFit` on the one advance model (`strAdvance` == the model behind `w`).
  const trailingSpaceW = snapToCharsClass(s, characterGrid)
    ? 0
    : s.text.endsWith(' ')
      ? w - strAdvance(s, trimmed)
      : 0;
  s.latinNaturalTrailingSpacePx =
    s.latinSpaceCompressionEligible === true && /^[^ ]+ $/u.test(s.text) && trailingSpaceW > 0
      ? trailingSpaceW
      : undefined;
  s.latinSpaceCompressionPx = undefined;
  const fitWidthFor = (
    widthPx: number,
    trailingSpacePx: number,
    next: LayoutSeg | undefined,
  ): number =>
    justifiedCandidateFitWidth(widthPx, trailingSpacePx, next, {
      isJustified,
      stretchLastLine,
      lineMaxWidth: breakerState.lineMaxWidth,
      lineXOffset: breakerState.lineXOffset,
      maxWidth,
    });
  const wForFit = fitWidthFor(prospectiveWidth, trailingSpaceW, breakerState.queue[0]);
  // ECMA-376 §17.3.1.33 does not prescribe a line-breaking tolerance.
  // Word-for-Mac controls with Calibri and Arial, left/center/right aligned
  // 10pt table cells, wrap a trailing Latin word below its natural advance
  // boundary (including <1pt overflow). Times New Roman differs by a
  // sub-point at that boundary, so this is a conservative library fit policy,
  // not a claim that every Office face and script has identical break points.
  // An earlier global 25%-of-spaces allowance pulled words up even when Word
  // did not; it also had no proven bound matching paint compression.
  // Dictionary-SEA candidate (Thai/Lao/Khmer; grapheme-fill Myanmar/Tibetan
  // stays on its per-cluster greedy path). Per-codepoint scan: a rare segment
  // mixing both SEA families is not dictionary-SEA, so
  // it keeps the pre-#991 greedy path instead of moving a grapheme-fill span
  // inside an atomic chunk.
  const sDictSea = s.seaBreaks !== undefined && isDictionarySeaText(s.text);

  // Atomic glued group: when THIS segment starts a glued group (its followers
  // in the queue are `joinPrev` pieces — small-caps case-pieces of the SAME
  // word like "I" then "NTRODUCTION", or a UAX#14 LB13 non-starter authored in
  // its own run like a trailing "," / "。"), the per-segment wrap below would
  // let the group split across lines. Pre-measure it and, if it does not fit on
  // the current (non-empty) line, flush so it starts fresh.
  //
  // ONLY when the lead segment is NOT itself CJK-breakable. A glued group whose
  // lead is a CJK run (e.g. "…通過する" + "。") is NOT atomic: the run splits at
  // an inter-CJK boundary and the trailing non-starter stays on its LAST piece
  // (§17.3.1.16 kinsoku keeps it off the next line's head when enabled — the
  // default; with kinsoku off it may lead the line, as it did before PR #602).
  // Pre-flushing the whole run instead leaves the prior line far short, which a
  // `both` line then stretches wide in a justified paragraph. `joinPrev`
  // stays a pure "this is a non-starter" marker; the atomic-vs-breakable decision lives here. A
  // non-breakable Latin / small-caps lead is genuinely atomic, so the pre-flush
  // (and the over-long-word char-break path below) still applies there.
  prepareAtomicTextFit(context, { s, w, trailingSpaceW, sDictSea, fitWidthFor });

  // §17.3.1.21 permits one eligible punctuation character past the text
  // extent. The isolated compatibility predicate owns both the CJK-language
  // sets and the bounded parent-run extensions owned by
  // `wordIsOverflowPunctuation`. CJK
  // segments that need an internal split retain their separate
  // overflowPunct-vs-kinsoku rule.
  const visibleSegmentScalars = [...trimmed];
  const trailingOverflowCharacter = visibleSegmentScalars.at(-1);
  const textBeforeTrailingOverflow = visibleSegmentScalars.slice(0, -1).join('');
  const admitsTrailingOverflowPunctuation =
    overflowPunct &&
    trailingOverflowCharacter !== undefined &&
    (breakerState.currentLine.length > 0 || textBeforeTrailingOverflow.length > 0) &&
    wordIsOverflowPunctuation(
      trailingOverflowCharacter,
      s.eastAsiaLanguage,
      s.overflowPunctuationEastAsianRun === true,
      s.script === 'ascii' || s.script === 'highAnsi',
      s.script === 'complexScript',
      s.overflowPunctuationBidiLanguage,
    ) &&
    breakerState.currentWidth + strAdvance(s, textBeforeTrailingOverflow) <= availW();

  // A line already admitted using this homogeneous-face rule cannot lend
  // that prior compression to a later mixed-face candidate. Its allocation
  // is finalized here, and the new route starts a fresh line.
  if (
    breakerState.latinAppliedPerGap > 0 &&
    (!breakerState.latinLineHomogeneous ||
      !breakerState.latinLineFace ||
      !sameLatinSpaceFace(s, breakerState.latinLineFace))
  ) {
    flush(undefined, false, s.src);
    breakerState.queue.unshift(s);
    return;
  }

  placeOrSplitText(context, {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  });
}

function processMathSegment(context: BreakOpportunityIteratorContext, seg: LayoutMathSeg): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    availW,
    setMeasureFont,
    fontFamilyClasses,
    measurement,
  } = context;

  const render = seg.mathMetadata;
  if (!render || render.available === false) {
    const emPx = seg.fontSize * scale;
    setMeasureFont(buildFont(false, false, emPx, null, fontFamilyClasses));
    const m = measurement.measureCurrentText(seg.fallbackText);
    const w = m.width;
    const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? emPx * 0.8;
    const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? emPx * 0.2;
    seg.measuredWidth = w;
    seg.mathAscent = asc;
    seg.mathDescent = desc;
    if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
      flush(undefined, false, seg.src);
    }
    addToLine(seg, w, seg.fontSize, Math.max(asc, emPx * 0.8), Math.max(desc, emPx * 0.2));
    return;
  }
  const emPx = seg.fontSize * scale;
  const w = render.widthEm * emPx;
  const asc = render.ascentEm * emPx;
  const desc = render.descentEm * emPx;
  seg.measuredWidth = w;
  // Ink extents (from the MathJax SVG viewBox) position the rasterized
  // glyph relative to the baseline when drawing.
  seg.mathAscent = asc;
  seg.mathDescent = desc;
  // …but the LINE BOX must reserve at least a normal single line for the
  // run's font size. A short equation — e.g. a lone "−" — has near-zero ink
  // height; using that as the line height would collapse the line (and the
  // table row) and pin the glyph to the very top of the cell. Floor to the
  // font's natural ascent/descent so math occupies a full line like text
  // does (tall math — fractions, big operators — keeps its larger ink box).
  const lineAsc = Math.max(asc, emPx * 0.8);
  const lineDesc = Math.max(desc, emPx * 0.2);
  if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
    flush(undefined, false, seg.src);
  }
  addToLine(seg, w, seg.fontSize, lineAsc, lineDesc);
  return;
}

function processImageSegment(context: BreakOpportunityIteratorContext, seg: LayoutImageSeg): void {
  const { breakerState, flush, addToLine, scale, availW } = context;

  if (seg.anchor) {
    seg.measuredWidth = 0;
    return;
  }
  const w = seg.widthPt * scale;
  const h = seg.heightPt;
  const asc = seg.heightPt * scale;
  seg.measuredWidth = w;
  if (breakerState.currentLine.length > 0 && breakerState.currentWidth + w > availW()) {
    flush(undefined, false, seg.src);
  }
  addToLine(seg, w, h, asc, 0);
  return;
}

function processTabSegment(context: BreakOpportunityIteratorContext, seg: LayoutTabSeg): void {
  const {
    breakerState,
    flush,
    baseRtl,
    addToLine,
    scale,
    firstIndent,
    tabOriginPx,
    maxWidth,
    marginRightPx,
    tabFollowWidth,
    measureText,
    verticalInkExtra,
    characterGrid,
    tabStops,
    defaultTabPt,
    tabFollowingMetrics,
    availW,
  } = context;

  // ── ECMA-376 §17.3.1.6 base-RTL ordinary tab ─────────────────────────
  // The LTR pen math below resolves stops in LOGICAL order, which mis-places
  // a bidi paragraph's tab-delimited cells (they reorder visually — see
  // `layoutBidiTabStops`). Add the tab with a PROVISIONAL width of 0 and do
  // NOT wrap on it; the per-line post-pass (`applyBidiTabs`, run in `flush`)
  // recomputes every tab width in the visual frame once the line's content
  // is known. A `<w:ptab>` (absolute-position tab) keeps the LTR path for
  // now (no bidi ptab fixture; its own NOTE flags the gap).
  if (baseRtl && !seg.ptab) {
    seg.measuredWidth = 0;
    addToLine(seg, 0, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    return;
  }

  // Absolute position on the line measured from paraX (line origin for continuation lines)
  const absFromParaX = breakerState.currentWidth + (breakerState.isFirst ? firstIndent : 0);

  // ── ECMA-376 §17.3.3.23 absolute-position tab (<w:ptab>) ──────────────
  // A ptab ignores the paragraph's custom tab stops and the default-tab
  // interval; it advances to a fixed position on the line derived from its
  // `alignment` (§17.18.71) and `relativeTo` (§17.18.73). The `alignment`
  // ALSO governs how the text after the ptab aligns to that position (left /
  // centered / right). All coordinates below are paraX-relative px.
  //
  // NOTE: the ptab target is resolved in LOGICAL (LTR) coordinates — this
  // block runs before the per-line bidi reorder pass, so it has no notion
  // of the paragraph's base direction. Interaction with bidi mirroring in
  // an RTL paragraph (where "left"/"right" alignment and the box edges
  // ought to mirror) is unverified; the primary use case (an LTR footer's
  // centered/right-aligned PAGE field) is correct.
  if (seg.ptab) {
    seg.resolvedAlignment = seg.ptab.alignment;
    // Reference box: "indent" ⇒ the paragraph content box [0, maxWidth];
    // "margin" ⇒ the text-margin box [-tabOriginPx, marginRightPx].
    const boxLeft = seg.ptab.relativeTo === 'indent' ? 0 : -tabOriginPx;
    const boxRight = seg.ptab.relativeTo === 'indent' ? maxWidth : marginRightPx;
    const target =
      seg.ptab.alignment === 'left'
        ? boxLeft
        : seg.ptab.alignment === 'center'
          ? (boxLeft + boxRight) / 2
          : boxRight;
    // Width of the content that trails the ptab up to the next tab / line end
    // — needed to right-/center-align it against `target` (the trailing text
    // is what aligns to the stop, §17.18.71).
    let followW = 0;
    for (const q of breakerState.queue) {
      if ('isTab' in q || 'lineBreak' in q) break;
      followW += tabFollowWidth(q);
    }
    const frac = seg.ptab.alignment === 'center' ? 0.5 : seg.ptab.alignment === 'right' ? 1 : 0;
    let tabW = target - absFromParaX - followW * frac;
    // §17.3.3.23: "If the alignment location … cannot be found on the current
    // line, because the starting location is past that point, then the tab …
    // shall advance to that location on the next available line." So when the
    // pen already sits at/after the target, wrap the ptab (and its trailing
    // content) to a fresh line — unless the line is empty (nowhere to wrap).
    if (tabW <= 0) {
      if (breakerState.currentLine.length > 0) {
        flush(undefined, false, seg.src);
        breakerState.queue.unshift(seg);
        return;
      }
      // Empty line: cannot advance backwards; contribute no width but keep the
      // segment so the line-height reflects the ptab's font.
      tabW = 0;
    }
    seg.measuredWidth = tabW;
    addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    // Commit the trailing content onto this line without a wrap re-check, so
    // it sits exactly at the aligned position (mirrors the custom right/center
    // tab path below).
    if (seg.ptab.alignment !== 'left') {
      while (breakerState.queue.length > 0) {
        const q = breakerState.queue[0];
        if ('isTab' in q || 'lineBreak' in q) break;
        breakerState.queue.shift();
        if ('imagePath' in q) {
          const w = q.widthPt * scale;
          q.measuredWidth = w;
          addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
        } else if ('math' in q) {
          addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
        } else {
          const m = measureText(q);
          // #1014 — fold the vo=Tr ink deficit into the committed advance too.
          const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), characterGrid, scale);
          q.measuredWidth = w;
          const asc =
            m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
          const desc =
            m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
          addToLine(q, w, q.fontSize, asc, desc);
        }
      }
    }
    return;
  }
  // ECMA-376 §17.3.1.37 / §17.15.1.25 — resolve the next stop in TEXT-MARGIN
  // coordinates (the same origin as custom stops): the current pen position is
  // `absFromParaX + tabOriginPx`, custom stops are `pos * scale`, and the
  // automatic grid interval is `defaultTabPt * scale`. Mixing paraX and margin
  // coordinates is what diverged leading-tab rows from labeled ones; computing
  // both in margin space and converting back keeps them aligned.
  const curMarginPx = absFromParaX + tabOriginPx;
  const customStopsPx = tabStops.map((t) => ({
    pos: t.pos * scale,
    alignment: t.alignment,
    leader: t.leader,
  }));
  const stop = nextTabStop(curMarginPx, customStopsPx, defaultTabPt * scale);
  seg.resolvedAlignment = stop?.alignment ?? 'left';
  // Convert the chosen margin-space stop back to paraX-relative px.
  const stopParaX = stop ? stop.pos - tabOriginPx : absFromParaX;
  // Right/center/decimal tab: place the tab + its trailing content (up to the next
  // tab / line end) so the content ends at / centers on the stop, and commit that
  // content directly so the normal wrap check doesn't push it past the stop
  // (ECMA-376 §17.3.1.37). This is what makes TOC "heading …… page" lines work.
  // Automatic stops returned by nextTabStop are left-aligned, so they fall
  // through to the left-tab path below.
  const alignmentRole = stop ? tabAlignmentRole(stop.alignment) : 'leading';
  if (stop && alignmentRole !== 'leading') {
    const stopX = stopParaX;
    seg.leader = stop.leader;
    const following = tabFollowingMetrics();
    const alignmentWidth =
      alignmentRole === 'center'
        ? following.totalWidth / 2
        : alignmentRole === 'decimal'
          ? (following.decimalPrefixWidth ?? following.totalWidth)
          : following.totalWidth;
    let tabW = stopX - absFromParaX - alignmentWidth;
    if (tabW <= 0) tabW = 0;
    seg.measuredWidth = tabW;
    addToLine(seg, tabW, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
    // Commit the trailing content onto this line without a wrap re-check.
    while (breakerState.queue.length > 0) {
      const q = breakerState.queue[0];
      if ('isTab' in q || 'lineBreak' in q) break;
      breakerState.queue.shift();
      if ('imagePath' in q) {
        const w = q.widthPt * scale;
        q.measuredWidth = w;
        addToLine(q, w, q.heightPt, q.heightPt * scale, 0);
      } else if ('math' in q) {
        addToLine(q, q.measuredWidth || 0, q.fontSize, q.mathAscent || 0, q.mathDescent || 0);
      } else {
        const m = measureText(q);
        // #1014 — fold the vo=Tr ink deficit into the committed advance too.
        const w = segAdvanceWidth(q, m.width + verticalInkExtra(q, q.text), characterGrid, scale);
        q.measuredWidth = w;
        const asc =
          m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? q.fontSize * scale * 0.8;
        const desc =
          m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? q.fontSize * scale * 0.2;
        addToLine(q, w, q.fontSize, asc, desc);
      }
    }
    return;
  }

  // Left-aligned tab (custom 'left'/'bar'/'clear' or an automatic stop): the
  // pen moves to the stop's paraX. nextTabStop already applied the §17.15.1.25
  // "after all custom stops" automatic grid, so there is no separate fallback.
  let tabWidth = stopParaX - absFromParaX;
  if (stop) seg.leader = stop.leader;
  // Clamp to avoid negative widths; if tab would overflow the line, wrap instead
  if (tabWidth <= 0) {
    flush(undefined, false, seg.src);
    breakerState.queue.unshift(seg);
    return;
  }
  if (breakerState.currentWidth + tabWidth > availW() && breakerState.currentLine.length > 0) {
    flush(undefined, false, seg.src);
    breakerState.queue.unshift(seg);
    return;
  }
  seg.measuredWidth = tabWidth;
  addToLine(seg, tabWidth, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
  return;
}

interface TextFitFrame {
  readonly s: LayoutTextSeg;
  readonly w: number;
  readonly h: number;
  readonly asc: number;
  readonly desc: number;
  readonly wForFit: number;
  readonly paragraphFinalIdeographicSpaceTail: boolean;
  readonly admitsTrailingOverflowPunctuation: boolean | undefined;
}

function placeOrSplitText(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    availW,
    fitsMeasuredWidth,
    fitHomogeneousLatinSpaces,
    appendQueuedIdeographicSpaceSegment,
    emergencyTextSplit,
    strNaturalAdvance,
    externalLinkSyntaxSplit,
    queueEmergencyTail,
  } = context;
  const {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;
  if (
    (fitsMeasuredWidth(breakerState.currentWidth + wForFit, availW()) &&
      breakerState.latinAppliedPerGap === 0) ||
    fitHomogeneousLatinSpaces(s, wForFit) ||
    fitsMeasuredWidth(breakerState.currentWidth + wForFit, availW())
  ) {
    // Fits on current line as-is
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    appendQueuedIdeographicSpaceSegment(s);
  } else if (admitsTrailingOverflowPunctuation) {
    s.measuredWidth = w;
    addToLine(s, w, h, asc, desc);
    appendQueuedIdeographicSpaceSegment(s);
  } else if (
    hasCJKBreakOpportunity(s.text) &&
    s.seaBreaks === undefined &&
    s.hardJoinPrev !== true
  ) {
    splitCjkOverflow(context, frame);
  } else if (s.seaBreaks !== undefined && s.hardJoinPrev !== true) {
    splitSeaOverflow(context, frame);
  } else if (breakerState.currentLine.length === 0) {
    // `word-overlong-token-emergency-break`: for a single non-CJK token wider
    // than a full line, fit the widest character prefix (at least one
    // character), draw it, and re-queue the remainder. Segments are already
    // space-delimited, so this cannot bypass an ordinary space opportunity.
    const split = externalLinkSyntaxSplit(s, availW()) || emergencyTextSplit(s, availW());
    if (split >= s.text.length) {
      // The visible glyphs actually fit (only a trailing space pushed it over the
      // fit test) — place the word whole.
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
    } else {
      const prefix = s.text.slice(0, split);
      const pw = strNaturalAdvance(s, prefix);
      addToLine(
        {
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        },
        pw,
        h,
        asc,
        desc,
      );
      queueEmergencyTail(s, split);
    }
  } else {
    const semanticSplit = externalLinkSyntaxSplit(s, availW() - breakerState.currentWidth);
    if (semanticSplit > 0 && semanticSplit < s.text.length) {
      const prefix = s.text.slice(0, semanticSplit);
      const pw = strNaturalAdvance(s, prefix);
      addToLine(
        {
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: prefix,
          measuredWidth: pw,
          ...slicedTextMetadata(s, 0, prefix.length),
        },
        pw,
        h,
        asc,
        desc,
      );
      queueEmergencyTail(s, semanticSplit);
      return;
    }
    if (s.joinPrev) {
      // LB14 and the other UAX glue rules prohibit a line boundary at this
      // source seam. If the complete glued group is wider than the fresh
      // line, split this member at the widest legal grapheme boundary that
      // fits the actual remaining band. This bases the decision on the group
      // advance, not on the follower's standalone width.
      const remaining = availW() - breakerState.currentWidth;
      const split = emergencyTextSplit(s, remaining, true);
      if ((remaining > 0 || s.hardJoinPrev === true) && split > 0 && split < s.text.length) {
        const prefix = s.text.slice(0, split);
        const pw = strNaturalAdvance(s, prefix);
        addToLine(
          {
            ...s,
            ...RESET_SLICED_TEXT_MEASUREMENT,
            text: prefix,
            measuredWidth: pw,
            ...slicedTextMetadata(s, 0, prefix.length),
          },
          pw,
          h,
          asc,
          desc,
        );
        queueEmergencyTail(s, split);
        return;
      }
      // A scalar span that continues the preceding grapheme (or another
      // explicitly glued piece) may overflow a pathological narrow line, but
      // it must never become a new line head and tear the cluster.
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
      return;
    }
    // Latin token does not fit on the current (non-empty) line: move it to a fresh
    // line and re-process. There it either fits, or — when it is wider than the
    // whole column — the empty-line branch above breaks it at the character level
    // (overflow-wrap). Re-queueing rather than force-adding is what lets that
    // over-long-word path run instead of letting the word spill the column.
    flush(undefined, false, s.src);
    breakerState.queue.unshift(s);
  }
}

function prepareAtomicTextFit(
  context: BreakOpportunityIteratorContext,
  frame: {
    s: LayoutTextSeg;
    w: number;
    trailingSpaceW: number;
    sDictSea: boolean;
    fitWidthFor: (widthPx: number, trailingSpacePx: number, next: LayoutSeg | undefined) => number;
  },
): void {
  const { breakerState, flush, availW, segAdvance, strAdvance } = context;
  const { s, w, trailingSpaceW, sDictSea, fitWidthFor } = frame;
  if (
    !s.joinPrev &&
    breakerState.currentLine.length > 0 &&
    (breakerState.queue[0] as LayoutTextSeg | undefined)?.joinPrev &&
    ((breakerState.queue[0] as LayoutTextSeg | undefined)?.hardJoinPrev === true ||
      !hasCJKBreakOpportunity(s.text)) &&
    // A SEA (Thai/Lao/Khmer) lead with usable word breaks is NOT atomic — the
    // run splits at a dictionary boundary (issue #797), mirroring the CJK gate.
    ((breakerState.queue[0] as LayoutTextSeg | undefined)?.hardJoinPrev === true ||
      !(s.seaBreaks && s.seaBreaks.length > 0))
  ) {
    let groupW = w;
    let groupTrail = trailingSpaceW;
    let groupEnd = 0;
    for (
      ;
      groupEnd < breakerState.queue.length &&
      (breakerState.queue[groupEnd] as LayoutTextSeg).joinPrev;
      groupEnd++
    ) {
      const f = breakerState.queue[groupEnd] as LayoutTextSeg;
      const hardPrefixEnd = hardJoinPrefixEnd(f);
      if (hardPrefixEnd !== undefined) {
        const prefix = f.text.slice(0, hardPrefixEnd);
        const prefixWidth = strAdvance(f, prefix);
        groupW += prefixWidth;
        groupTrail = prefix.endsWith(' ')
          ? prefixWidth - strAdvance(f, prefix.replace(/ +$/, ''))
          : 0;
        // A whole hard member can lead into another hard member. Otherwise
        // the first legal boundary after the seam ends the atomic prefix.
        if (hardPrefixEnd < f.text.length) break;
        continue;
      }
      const firstExternalBreak = f.externalLinkBreakOffsets?.[0];
      if (firstExternalBreak !== undefined) {
        const prefix = f.text.slice(0, firstExternalBreak);
        const prefixWidth = strAdvance(f, prefix);
        groupW += prefixWidth;
        groupTrail = 0;
        break;
      }
      // A CJK-BREAKABLE follower (e.g. "Roman" + "、あるいは…用いる。") is NOT
      // atomic: only its LEADING run of line-start-forbidden chars would orphan
      // at a line head (UAX#14 LB13 / §17.3.1.16); the rest splits at an
      // inter-CJK boundary and wraps on its own. So glue only that prefix's
      // advance to the lead and STOP summing here — mirror of the CJK-lead
      // direction handled by the `!hasCJKBreakOpportunity(s.text)` gate above.
      // Summing the whole breakable run would pre-flush the Latin lead alone,
      // leaving a `both` line stretched sparse. A Latin / small-caps follower (no CJK break opportunity —
      // the "I" + "NTRODUCTION" case) stays fully atomic: keep full-add.
      if (hasCJKBreakOpportunity(f.text)) {
        const chars = [...f.text];
        let p = 0;
        while (
          p < chars.length &&
          DEFAULT_KINSOKU_RULES.lineStartForbidden.has(chars[p].codePointAt(0)!)
        )
          p++;
        if (p < chars.length) {
          // Breakable rest exists past the leading non-starters: glue only the
          // prefix (it may be empty — then "Roman" is effectively unglued and
          // wraps on its own) and end the atomic group here.
          const prefix = chars.slice(0, p).join('');
          const prefixWidth = strAdvance(f, prefix);
          groupW += prefixWidth;
          groupTrail = 0;
          break;
        }
        // Entirely non-starters (no breakable rest): fall through to full-add.
      }
      const fw = segAdvance(f);
      groupW += fw;
      const ft = f.text.replace(/ +$/, '');
      const followerTrail = f.text.endsWith(' ') ? fw - strAdvance(f, ft) : 0;
      // UAX #14 LB7 makes a consecutive SP sequence one trailing suffix even
      // when a source-formatting boundary split it into multiple segments.
      // Accumulate space-only followers so the line-end fit allowance is
      // invariant to that non-textual boundary. A follower containing visible
      // text starts a new suffix and therefore replaces the previous value.
      groupTrail = ft.length === 0 && groupTrail > 0 ? groupTrail + followerTrail : followerTrail;
    }
    if (
      breakerState.currentWidth + fitWidthFor(groupW, groupTrail, breakerState.queue[groupEnd]) >
      availW()
    ) {
      flush(undefined, false, s.src);
    }
  }

  // `word-dictionary-sea-atomic-chunk`: ECMA-376 prescribes no SEA
  // line-breaking algorithm. Treat dictionary boundaries inside a no-space
  // Thai/Lao/Khmer chunk as secondary opportunities: move a chunk that fits a
  // full line as a unit; only a full-line-overlong chunk breaks at dictionary
  // boundaries through the greedy SEA branch below.
  //
  // Judged only at chunk START: if the previously committed token is a text
  // segment glued to `s` (no trailing space), the whole chunk already passed
  // this judgment when its head was placed, so a mid-chunk segment never
  // needs it. The chunk spans `s` plus following queue segments while they
  // stay dictionary-SEA text glued without intervening spaces. Grapheme-fill
  // scripts (Myanmar/Tibetan) are excluded because their per-cluster path
  // fills the remaining width.
  if (
    sDictSea &&
    breakerState.currentLine.length > 0 &&
    (() => {
      const last = breakerState.currentLine[breakerState.currentLine.length - 1];
      return !('text' in last) || (last as LayoutTextSeg).text.endsWith(' ');
    })()
  ) {
    let chunkW = w;
    let chunkTrail = trailingSpaceW;
    let chunkEnd = 0;
    if (!s.text.endsWith(' ')) {
      for (; chunkEnd < breakerState.queue.length; chunkEnd++) {
        const f = breakerState.queue[chunkEnd];
        if (!('text' in f) || (f as LayoutTextSeg).seaBreaks === undefined) break;
        if (!isDictionarySeaText((f as LayoutTextSeg).text)) break;
        const ft = f as LayoutTextSeg;
        const fw = segAdvance(ft);
        const fTrim = ft.text.replace(/ +$/, '');
        chunkW += fw;
        chunkTrail = ft.text.endsWith(' ') ? fw - strAdvance(ft, fTrim) : 0;
        if (ft.text.endsWith(' ')) {
          chunkEnd++;
          break;
        } // a space ends the chunk
      }
    }
    const chunkWForFit = fitWidthFor(chunkW, chunkTrail, breakerState.queue[chunkEnd]);
    if (
      breakerState.currentWidth + chunkWForFit > availW() &&
      chunkWForFit <= breakerState.lineMaxWidth
    ) {
      flush(undefined, false, s.src);
    }
  }
}

function splitCjkOverflow(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    characterGrid,
    availW,
    setMeasureFont,
    fontFamilyClasses,
    measurement,
    strAdvance,
    overflowPunct,
    appendQueuedIdeographicSpaceSegment,
    emergencyTextSplit,
    effectiveFontPx,
    ctx,
    verticalGlyphMeasurement,
    kinsoku,
    strNaturalAdvance,
    retractCurrentLineForLeadingKinsoku,
    keepLeadingKinsokuWithCurrentLine,
  } = context;
  const {
    s,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;

  // CJK overflow: split at the maximum prefix that fits, re-queue the tail.
  // A segment that ALSO contains SEA (a mixed CJK+SEA `<w:cs/>` run) is routed
  // to the SEA branch below instead — its `seaBreaks` already merges the CJK
  // per-character opportunities with the SEA dictionary/transition ones
  // (issue #960), so both scripts break by their own rule from one offset set.
  // (pptx's analogous CJK fit is cjk-wrap.ts `fitCjkLine`, kept intentionally
  //  separate: it sums per-char advances, whereas this path uses substring
  //  binary-search + the cross-run 追い出し below. Don't naively unify them.)
  const available = availW() - breakerState.currentWidth;
  let rawPrefix = '';
  const maximumIdeographicSpaceHang = paragraphFinalIdeographicSpaceTail
    ? wordIdeographicSpaceLineEndAllowanceCount(
        hasEastAsianVisiblePredecessor(s.text),
        s.paragraphFinalIdeographicSpaceCount ?? 0,
      )
    : Number.POSITIVE_INFINITY;
  if (available > 0) {
    const nonMonotoneAllocation =
      charSpacingDeltaPx(s, scale) < 0 || snapToCharsClass(s, characterGrid) === 'latin';
    if (nonMonotoneAllocation) {
      rawPrefix = s.text.slice(0, emergencyTextSplit(s, available, false));
    } else {
      setMeasureFont(
        buildFont(
          s.bold,
          s.italic,
          effectiveFontPx(s),
          s.fontFamily,
          fontFamilyClasses,
          s.fontRoute,
        ),
      );
      measurement.withSegmentKerning(s, () => {
        rawPrefix = fitCJKPrefix(
          ctx,
          s.text,
          available,
          segmentCharacterGridDeltaPx(s, characterGrid, scale),
          charScaleFactor(s),
          charSpacingDeltaPx(s, scale),
          s.verticalRun === true,
          verticalGlyphMeasurement,
          (prefix) => strAdvance(s, prefix),
          maximumIdeographicSpaceHang,
        );
      });
    }
  }
  // Apply kinsoku to the break position: retract leftwards so the tail
  // never begins with a 行頭禁則 char and the head never ends with a
  // 行末禁則 char (ECMA-376 §17.15.1.58–.60). When the current line
  // already has content, retracting to an empty prefix is allowed — the
  // whole run moves to the next (fresh) line, which is Word's 追い出し.
  // When the line is empty we keep at least one char (minSplit=1) so we
  // never lose forward progress.
  const allChars = [...s.text];
  const rawSplit = [...rawPrefix].length;
  const minSplit = breakerState.currentLine.length > 0 ? 0 : 1;
  // ECMA-376 §17.3.1.21 permits one punctuation character beyond the
  // paragraph extents. The isolated compatibility projection resolves the
  // language-specific set and its precedence over kinsoku at this internal
  // CJK split.
  const hangingSplit =
    overflowPunct &&
    rawSplit < allChars.length &&
    (breakerState.currentLine.length > 0 || rawSplit > 0) &&
    wordIsOverflowPunctuation(
      allChars[rawSplit],
      s.eastAsiaLanguage,
      s.overflowPunctuationEastAsianRun === true,
      s.script === 'ascii' || s.script === 'highAnsi',
      s.script === 'complexScript',
      s.overflowPunctuationBidiLanguage,
    )
      ? rawSplit + 1
      : null;
  const proposedSplit = extendThroughTrailingIdeographicSpaces(
    allChars,
    hangingSplit ?? kinsokuAdjustedSplit(allChars, rawSplit, kinsoku, minSplit),
    paragraphFinalIdeographicSpaceTail && maximumIdeographicSpaceHang === 0
      ? 0
      : maximumIdeographicSpaceHang,
  );
  const proposedPrefix = allChars.slice(0, proposedSplit).join('').length;
  const protectedSplit = legalTextSplitAtOrBefore(s, proposedPrefix, minSplit > 0 ? 1 : 0);
  const prefix = s.text.slice(0, protectedSplit);
  if (prefix.length > 0) {
    // Grid advance for the head piece — the same model as the line box / draw.
    const pw = strNaturalAdvance(s, prefix);
    const headSeg: LayoutTextSeg = {
      ...s,
      ...RESET_SLICED_TEXT_MEASUREMENT,
      text: prefix,
      measuredWidth: pw,
      ...slicedTextMetadata(s, 0, prefix.length),
    };
    addToLine(headSeg, pw, h, asc, desc);
    const tail = s.text.slice(prefix.length);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, prefix.length, s.text.length),
        measuredWidth: 0,
        src: {
          segIndex: s.src!.segIndex,
          charOffset: s.src!.charOffset + prefix.length,
        },
      });
    } else {
      appendQueuedIdeographicSpaceSegment(s);
    }
  } else if (breakerState.currentLine.length > 0) {
    // No prefix of `s` fits. If `s` would lead the next line with a 行頭禁則
    // char, kinsokuAdjustedSplit can't fix it from within `s` (the offending
    // char is its first); pull trailing graphemes of the current line's last
    // text segment down so they lead the next line ahead of `s` — cross-run
    // 追い出し (§17.3.1.16). See crossRunKinsokuRetract for the bounded,
    // re-validating, whitespace-guarded retraction count.
    const retraction = retractCurrentLineForLeadingKinsoku(s);
    if (retraction.kind === 'blocked') {
      keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
      return;
    }
    flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
    breakerState.queue.unshift(s);
    if (retraction.kind === 'retracted') breakerState.queue.unshift(retraction.tail);
  } else {
    // Empty line and not even one char fits — force-fit one char to guarantee progress
    const forcedChars = [...s.text];
    const forcedSplit =
      forcedChars.length > 0
        ? extendThroughTrailingIdeographicSpaces(
            forcedChars,
            1,
            s.paragraphFinalIdeographicSpaceTail === true
              ? wordIdeographicSpaceLineEndAllowanceCount(
                  EAST_ASIAN_RE.test(forcedChars[0] ?? ''),
                  s.paragraphFinalIdeographicSpaceCount ?? 0,
                )
              : Number.POSITIVE_INFINITY,
          )
        : 0;
    const forcedUtf16 = forcedChars.slice(0, forcedSplit).join('').length;
    const legalForcedUtf16 =
      legalTextSplitAtOrBefore(s, forcedUtf16) || emergencyTextSplit(s, availW(), true);
    const firstChar = s.text.slice(0, legalForcedUtf16);
    if (firstChar) {
      const fw = strNaturalAdvance(s, firstChar);
      const headSeg: LayoutTextSeg = {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: firstChar,
        measuredWidth: fw,
        ...slicedTextMetadata(s, 0, firstChar.length),
      };
      addToLine(headSeg, fw, h, asc, desc);
      const tail = s.text.slice(firstChar.length);
      if (tail) {
        breakerState.queue.unshift({
          ...s,
          ...RESET_SLICED_TEXT_MEASUREMENT,
          text: tail,
          ...slicedTextMetadata(s, firstChar.length, s.text.length),
          measuredWidth: 0,
          src: {
            segIndex: s.src!.segIndex,
            charOffset: s.src!.charOffset + firstChar.length,
          },
        });
      } else {
        appendQueuedIdeographicSpaceSegment(s);
      }
    }
  }
}

function splitSeaOverflow(context: BreakOpportunityIteratorContext, frame: TextFitFrame): void {
  const {
    breakerState,
    flush,
    addToLine,
    scale,
    characterGrid,
    availW,
    strAdvance,
    emergencyTextSplit,
    strNaturalAdvance,
    retractCurrentLineForLeadingKinsoku,
    keepLeadingKinsokuWithCurrentLine,
  } = context;
  const {
    s: segment,
    w,
    h,
    asc,
    desc,
    wForFit,
    paragraphFinalIdeographicSpaceTail,
    admitsTrailingOverflowPunctuation,
  } = frame;
  const s = segment as LayoutTextSeg & { seaBreaks: readonly number[] };

  // No-inter-word-space line wrap: Thai/Lao/Khmer dictionary words (#797) or
  // Myanmar/Tibetan grapheme clusters (#961). This ONE segment is a whole run;
  // break it only at a member of `s.seaBreaks` — the UNION (#960) of the
  // dictionary word (or grapheme-cluster) boundaries, the no-space SEA↔non-SEA
  // script transitions, and (for a mixed CJK+SEA `<w:cs/>` run) the CJK
  // per-character opportunities, already kinsoku-filtered by
  // `seaMixedBreakOffsets`. Entered for ANY such segment (even one with no
  // interior boundary — a single word/cluster wider than the column, or
  // Segmenter unavailable) so the emergency split below stays GRAPHEME-safe
  // instead of falling to the code-point path. Kinsoku 行頭/行末禁則 was applied
  // when the offsets were built (so a forbidden CJK char never heads/tails a
  // line); choosing an earlier legal offset is the only remaining adjustment,
  // which fitSeaWordPrefix already does. The run stays one contiguous draw per
  // line (measure==paint); the tail re-queues with its offsets rebased.
  const available = availW() - breakerState.currentWidth;
  const measureSub = (sub: string): number => strAdvance(s, sub);
  // Grapheme-fill runs (Myanmar/Tibetan) have DENSE offsets (one per cluster),
  // so use the monotone binary-search fit — a per-line full scan would be O(n²)
  // down a long run. Dictionary runs keep the negative-spacing-safe full scan.
  const monotone =
    isGraphemeFillText(s.text) &&
    charSpacingDeltaPx(s, scale) >= 0 &&
    snapToCharsClass(s, characterGrid) !== 'latin';
  const split = fitSeaWordPrefix(s.text, s.seaBreaks, 0, available, measureSub, monotone);
  if (split > 0) {
    const prefix = s.text.slice(0, split);
    const pw = strNaturalAdvance(s, prefix);
    addToLine(
      {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: prefix,
        measuredWidth: pw,
        ...slicedTextMetadata(s, 0, prefix.length),
      },
      pw,
      h,
      asc,
      desc,
    );
    const tail = s.text.slice(split);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, split, s.text.length),
        measuredWidth: 0,
        src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + split },
        seaBreaks: rebaseSeaBreaks(s.seaBreaks, split),
      });
    }
  } else if (breakerState.currentLine.length > 0) {
    // No whole word fits the remaining band — move the run to a fresh line and
    // re-process (Latin-word style). If `s` would then LEAD the next line with
    // a 行頭禁則 char (a mixed CJK+SEA run whose first glyph is a forbidden
    // leader — #960 routes it here, where the offset set cannot fix a
    // segment-initial char), pull trailing graphemes of the current line's
    // last text segment down so they lead ahead of `s` — the same cross-run
    // 追い出し (§17.3.1.16) the CJK branch does.
    const retraction = retractCurrentLineForLeadingKinsoku(s);
    if (retraction.kind === 'blocked') {
      keepLeadingKinsokuWithCurrentLine(s, h, asc, desc);
      return;
    }
    flush(undefined, false, retraction.kind === 'retracted' ? retraction.tail.src : s.src);
    breakerState.queue.unshift(s);
    if (retraction.kind === 'retracted') breakerState.queue.unshift(retraction.tail);
  } else {
    // Empty line and the first dictionary word is wider than the whole
    // column: emergency GRAPHEME-safe split (a code-point split would tear a
    // base + tone/combining mark, both BMP). Guarantee ≥1 cluster of progress.
    const firstWordEnd = s.seaBreaks[0] ?? s.text.length;
    const firstWord = s.text.slice(0, firstWordEnd);
    const graphemes = graphemeClusterOffsets(firstWord);
    let gsplit = fitSeaWordPrefix(firstWord, graphemes, 0, available, measureSub, monotone);
    if (gsplit <= 0) gsplit = graphemes.length > 0 ? graphemes[0] : firstWord.length;
    gsplit = legalTextSplitAtOrBefore(s, gsplit) || emergencyTextSplit(s, available, true);
    const prefix = s.text.slice(0, gsplit);
    const pw = strNaturalAdvance(s, prefix);
    addToLine(
      {
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: prefix,
        measuredWidth: pw,
        ...slicedTextMetadata(s, 0, prefix.length),
      },
      pw,
      h,
      asc,
      desc,
    );
    const tail = s.text.slice(gsplit);
    if (tail) {
      breakerState.queue.unshift({
        ...s,
        ...RESET_SLICED_TEXT_MEASUREMENT,
        text: tail,
        ...slicedTextMetadata(s, gsplit, s.text.length),
        measuredWidth: 0,
        src: { segIndex: s.src!.segIndex, charOffset: s.src!.charOffset + gsplit },
        seaBreaks: rebaseSeaBreaks(s.seaBreaks, gsplit),
      });
    }
  }
}
