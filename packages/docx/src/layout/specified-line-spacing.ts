import type { LayoutLine, LayoutTextSeg } from '../line-layout.js';
import type { NumberingMarkerShapeInput } from './types.js';
import type { ParagraphLayoutSource, ParagraphTextBearingRun } from './text.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';
import { wordSpecifiedTextLineMetrics } from './line-compatibility.js';

function referenceSourceAdmitted(segment: LayoutTextSeg): boolean {
  return segment.fontSource === 'native'
    || (segment.fontSource === 'local' && segment.authoredReferenceMetricAllowed === true)
    || (segment.fontSource === undefined && segment.fontRoute?.scope === 'native');
}

/** Check the complete paragraph once per allocation/retention pass. A later
 * decorated or non-Latin run must not switch an earlier continuation line to
 * a different baseline contract. Separate mixed-script/decorated paragraphs
 * do not establish the homogeneous matrix's rule, even on their ASCII-only
 * physical lines. Keeping admission paragraph-wide also avoids quadratic
 * rescans of source runs for long wrapped paragraphs. */
export function specifiedTextParagraphIsHomogeneous(paragraph: ParagraphLayoutSource): boolean {
  return paragraph.runs.every(run => {
    if (run.type === 'field' || run.type === 'math') return false;
    if (run.type !== 'text') return true;
    if (!/^[\x20-\x7e]*$/u.test(run.text) || run.bold || run.italic
      || run.underline || run.strikethrough || run.hyperlink || run.vertAlign
      || run.position || run.smallCaps || run.allCaps || run.ruby
      || run.charSpacing || (run.charScale !== undefined && run.charScale !== 100)
      || (run.fontFamily ?? paragraph.defaultFontFamily) !== paragraph.defaultFontFamily
      || run.fontSize !== paragraph.defaultFontSize) return false;
    const fontAxes: ParagraphTextBearingRun = run;
    for (const axis of ['fontFamilyHighAnsi', 'fontFamilyEastAsia', 'fontFamilyCs'] as const) {
      if (fontAxes[axis] && fontAxes[axis] !== paragraph.defaultFontFamily) return false;
    }
    return true;
  });
}

/** Evidence gate for the ordinary Latin class in WORD_SPECIFIED_TEXT_LINE_BOX.
 * Other scripts, mixed faces/sizes, transformed runs, inline objects, ruby,
 * numbering, decorated text, differing paragraph marks and non-native resources
 * retain their existing allocation. The authored reference is metadata policy,
 * not proof of an installed paint face. A positively loaded local FontFace
 * installed alias is the same authored-face class as the native route; its selected
 * source and authored name come from the resolver, never CSS-name inference.
 * The existing authored-reference predicate admits installed Office local()
 * tuples, not application-provided SFNT resources registered as local faces.
 * Separate Office-PDF comparisons of decorated exact paragraphs did not
 * support extending the fixed partition to their mark/style mixtures. Those
 * counterexamples do not isolate decoration from mark formatting: preserve
 * both unmeasured classes instead of assigning a guessed correction. */
export function specifiedTextLineMetrics(
  line: LayoutLine,
  context: ParagraphLayoutContext,
  mark: Pick<ParagraphLayoutSource, 'defaultFontSize' | 'defaultFontFamily' | 'paragraphMarkShapeInput'>,
  mode: number | undefined,
  vertical = false,
  markShape?: NumberingMarkerShapeInput,
): Readonly<{ advancePt: number; baselineOffsetPt: number }> | null {
  const spacing = context.lineSpacing;
  if ((mode !== 14 && mode !== 15) || vertical || context.hasRuby
    || spacing?.explicit !== true || spacing.value <= 0
    || (spacing.rule !== 'exact' && spacing.rule !== 'atLeast'
      && !(spacing.rule === 'auto' && spacing.value === 1))) return null;
  // The independent matrix establishes exact heights from 12 through 30pt.
  // Separate smaller-line PDF captions accumulate a device cadence that does
  // not isolate their baseline partition from preceding advances. Preserve
  // heights outside the measured interval; these are evidence bounds, not a
  // fitted minimum line height or a correction to the authored spacing.
  if (spacing.rule === 'exact' && (spacing.value < 12 || spacing.value > 30)) return null;
  // No-grid auto=1 diagnostics already agree with the established normal-box
  // path. Only the measured whole-cell grid diagnostic needs this override.
  if (spacing.rule === 'auto' && !context.lineGrid.active) return null;
  // This round's hosts have wrap=none. A separate narrow-side-gap Office
  // counterexample changes the physical partition and its exclusion-conditioned
  // origin together, so it cannot establish an independent baseline rule.
  // topY is supplied only by the active wrap oracle, including exploratory
  // allocation passes before wrapAllocation is published. Keep that complete
  // class on its existing convergence/placement contract.
  if (line.topY !== undefined) return null;
  // Grid-atLeast is deliberately unresolved: the printing PDFs keep anchors
  // on exact point advances while text can follow a different device cadence
  // (40pt minima produced successive 40.08pt text steps on a 20pt grid). Neither
  // §17.3.1.33 nor §17.6.5 establishes that projection. Preserve the complete
  // existing class rather than fit the small baseline residual or its hosts.
  if (spacing.rule === 'atLeast' && context.lineGrid.active) return null;
  let first: LayoutTextSeg | undefined;
  for (const segment of line.segments) {
    if (!('text' in segment)) return null;
    if (segment.metricOnly) continue;
    if (!segment.text) continue;
    if (!/^[\x20-\x7e]+$/u.test(segment.text) || segment.bold || segment.italic
      || segment.underline || segment.strikethrough || segment.hyperlink
      || segment.ruby || segment.vertAlign || segment.position || segment.smallCaps
      || segment.verticalRun || segment.textBoxLineFloor
      || !referenceSourceAdmitted(segment)) return null;
    first ??= segment;
    if ((segment.authoredFontFamily ?? segment.fontFamily) !== (first.authoredFontFamily ?? first.fontFamily)
      || segment.fontSize !== first.fontSize
      || segment.fontRoute?.fingerprint !== first.fontRoute?.fingerprint) return null;
  }
  // Body acquisition retains the full mark axes on the source paragraph;
  // shape-text callers may also supply an explicit acquisition input. The
  // selected Latin mark face alone cannot prove homogeneous script axes.
  const paragraphMark = mark.paragraphMarkShapeInput ?? markShape;
  if (paragraphMark && (paragraphMark.weight !== 400 || paragraphMark.style !== 'normal'
    || paragraphMark.complexScript)) return null;
  const authoredFamily = first?.authoredFontFamily ?? first?.fontFamily;
  if (!first || mark.defaultFontFamily !== authoredFamily
    || mark.defaultFontSize !== first.fontSize || first.fontSize <= 0
    || !Number.isFinite(first.fontSize)) return null;
  if (paragraphMark) {
    if (paragraphMark.fontSizePt !== first.fontSize) return null;
    for (const family of Object.values(paragraphMark.fonts)) {
      if (family && family !== authoredFamily) return null;
    }
  }
  // Host-face/style changes were not part of the homogeneous controls.
  for (const segment of line.segments) {
    if ('text' in segment && segment.metricOnly
      && ((segment.authoredFontFamily ?? segment.fontFamily) !== authoredFamily
        || segment.bold || segment.italic || !referenceSourceAdmitted(segment))) return null;
  }
  const profile = referenceFontLineMetrics(authoredFamily, 400, 'normal');
  if (!profile) return null;
  return wordSpecifiedTextLineMetrics({
    rule: spacing.rule,
    value: spacing.value,
    descentPt: profile.designDescentRatio * first.fontSize,
    singlePt: profile.lineHeightRatio * first.fontSize,
    pitchPt: context.lineGrid.active ? context.lineGrid.pitchPt : null,
  });
}
