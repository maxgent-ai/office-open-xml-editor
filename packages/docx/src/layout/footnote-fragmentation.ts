import { sliceParagraphLayout } from './paragraph.js';
import { projectBodyOccurrence } from './occurrence-projection.js';
import { isDeepFrozenPlainDataRoot } from './plain-data.js';
import { placeNoteSeparatorOccurrence } from './native-note-separators.js';
import { unionLayoutRects } from './rect-union.js';
import type { NoteLayout, NoteSeparatorLayout, PaintNode, ParagraphLayout, StoryLayout } from './types.js';

export interface FootnoteCursor {
  readonly blockIndex: number;
  readonly lineIndex: number;
  readonly inlineExtentPt: number;
  readonly sourcePosition: Readonly<{ runIndex: number; offset: number }> | number | null;
}

/** Resolve a shaped cut back to its source run, independent of preceding
 * PAGE/NUMPAGES display lengths. A line cursor may be reused only when its
 * source cut is unchanged. Source-offset reflow is separate layout work; this
 * guard prevents repeated or omitted text while that work remains unresolved. */
interface ParagraphContinuationIndex {
  readonly advances: Float64Array;
  readonly runStarts: ReadonlyMap<number, number>;
}

function indexParagraph(block: ParagraphLayout): ParagraphContinuationIndex {
  const advances = new Float64Array(block.lines.length + 1);
  const runStarts = new Map<number, number>();
  const record = (runIndex: number, start: number) =>
    runStarts.set(runIndex, Math.min(runStarts.get(runIndex) ?? start, start));
  block.lines.forEach((line, index) => {
    const previous = block.lines[index - 1];
    const gap = previous ? Math.max(0, line.bounds.yPt - (previous.bounds.yPt + previous.advancePt)) : 0;
    advances[index + 1] = advances[index]! + gap + line.advancePt;
    if (!Number.isFinite(advances[index + 1])) throw new RangeError('Invalid footnote line advance');
    for (const placement of line.placements) {
      if (placement.kind === 'text' && placement.sourceRuns) {
        for (const run of placement.sourceRuns) record(run.sourceRunIndex, run.range.start);
      } else if ((placement.kind === 'text' || placement.kind === 'resource')
        && placement.sourceRunIndex !== undefined) record(placement.sourceRunIndex, placement.range.start);
    }
  });
  return { advances, runStarts };
}

function lineSourcePosition(block: ParagraphLayout, lineIndex: number,
  index: ParagraphContinuationIndex): FootnoteCursor['sourcePosition'] {
  const line = block.lines[lineIndex];
  if (!line) return null;
  const first = line.placements.find(placement => placement.kind !== 'text'
    || placement.role !== 'numbering-marker');
  // Coalesced glyph sequences retain each original run's ownership range.
  // sourceRunIndex alone names only the sequence's first run and would reject
  // an unchanged cut whenever an earlier resolved field changes display length.
  const owner = first?.kind === 'text'
    ? first.sourceRuns?.find(run => run.range.start <= first.range.start && first.range.start < run.range.end)
    : undefined;
  const runIndex = owner?.sourceRunIndex ?? (first && (first.kind === 'text' || first.kind === 'resource')
    ? first.sourceRunIndex : undefined);
  if (first && runIndex !== undefined) {
    const runStart = index.runStarts.get(runIndex) ?? first.range.start;
    return Object.freeze({ runIndex, offset: first.range.start - runStart });
  }
  // Atomic controls without source-run metadata use their occurrence offset.
  // A changed offset is rejected conservatively rather than guessing a cut.
  return line.range.start;
}

function sameSourcePosition(a: FootnoteCursor['sourcePosition'], b: FootnoteCursor['sourcePosition']): boolean {
  return typeof a === 'object' && a !== null && typeof b === 'object' && b !== null
    ? a.runIndex === b.runIndex && a.offset === b.offset : a === b;
}

/**
 * ECMA-376 §17.11.1 illustrates a footnote continued across pages, and
 * §17.11.21 locates its band at the page bottom. Once pagination admits a
 * continuation, library policy retains complete shaped lines within capacity. The following page owns a
 * fresh note separator and resumes at the first unpainted line. A note whose
 * first line cannot fit is left to ordinary reference relocation.
 */
export function partitionFootnote(
  acquired: NoteLayout,
  cursor: FootnoteCursor | null,
  capacityPt: number,
  continuationNotice?: ContinuationNoticeProvider,
  extent: FootnotePartitionExtent = 'largest',
): FootnotePartition | null {
  return createFootnotePartitioner()(acquired, cursor, capacityPt, continuationNotice, extent);
}

/**
 * A retained fragment, the cursor its note resumes at, and `hostShiftPt`:
 * the y translation the source-cut projection gives the host flow (lines and
 * host-following drawings) of the fragment's first retained paragraph whose
 * text box story is band dependent, from the acquisition's story coordinates
 * to the fragment's; 0 without one. The projection keeps the acquisition's
 * cursor deltas, so every retained block receives that translation. A head
 * starts at its story's origin (0); a continued tail moves up by its cut.
 * Such a story places page-owned content against the band its acquisition is
 * given (body-paginator.ts footnoteBandPlan), so that band must include it.
 */
export type FootnotePartition = Readonly<{
  fragment: NoteLayout;
  nextCursor: FootnoteCursor | null;
  hostShiftPt: number;
}>;

/** `largest` keeps every complete line that fits the capacity (admission);
 * a continuing fragment reserves the page's notice inside that capacity.
 * `largest-with-reserved-notice` is the same selection when the caller has
 * already reserved the page's single notice outside the capacity (because a
 * later note of the same body unit may continue): a continuing fragment
 * still carries the notice occurrence but is not charged it a second time.
 * `minimum` keeps only the first real line (plus any leading empty
 * paragraphs) regardless of capacity and reserves no notice: the legal
 * minimum head that pagination must leave for a later reference in the same
 * mandatory body unit. A note that cannot partition still returns null. */
export type FootnotePartitionExtent = 'largest' | 'largest-with-reserved-notice' | 'minimum';

/** Lazily acquired, unplaced native continuation notice occurrence; null when
 * the story is absent, empty, guard-only or fully hidden. */
export type ContinuationNoticeProvider = () => NoteSeparatorLayout | null;

/** One paginator pass owns these indexes. Weak keys do not retain an expired
 * acquisition; scalar prefix/range data does not point back to its key. Reuse
 * only engine-sealed roots, never arbitrary shallow-frozen/mutable callers.
 * Admission uses cumulative sums, but final fragment advance keeps the shared
 * paragraph finalizer's addition order. Page-local reacquisition gets a new key. */
export function createFootnotePartitioner(): typeof partitionFootnote {
  const indexes = new WeakMap<ParagraphLayout, ParagraphContinuationIndex>();
  const supportedStories = new WeakMap<StoryLayout, boolean>();
  const indexOf = (block: ParagraphLayout) => {
    let index = indexes.get(block);
    if (!index) {
      index = indexParagraph(block);
      if (isDeepFrozenPlainDataRoot(block)) indexes.set(block, index);
    }
    return index;
  };
  const supports = (story: StoryLayout) => {
    let supported = supportedStories.get(story);
    if (supported === undefined) {
      supported = story.blocks.every(block => block.kind === 'paragraph' && block.ordinaryFlow);
      if (isDeepFrozenPlainDataRoot(story)) supportedStories.set(story, supported);
    }
    return supported;
  };
  return (acquired, cursor, capacityPt, continuationNotice, extent = 'largest') => {
    if (acquired.trailing) throw new Error('An acquired note cannot own a continuation notice');
    if (extent === 'minimum') {
      return partitionIndexedFootnote(acquired, cursor, Number.MAX_VALUE, indexOf, supports, 1);
    }
    const partition = partitionIndexedFootnote(acquired, cursor, capacityPt, indexOf, supports);
    if (!partition?.nextCursor || !continuationNotice) return partition;
    const notice = continuationNotice();
    if (!notice) return partition;
    // ECMA-376 §17.18.33 places the notice at the bottom of a page on which
    // the note continues, so only an actually continuing fragment reserves
    // its measured advance. One further partition with that reserve is final:
    // a smaller capacity cannot complete the note, and if no real note line
    // remains the reference relocates (or admission reports capacity) rather
    // than retaining a separator/notice-only fragment. When the caller has
    // already reserved the notice outside the capacity, the first partition
    // is final and only carries the occurrence.
    const reserved = extent === 'largest-with-reserved-notice'
      ? partition
      : partitionIndexedFootnote(acquired, cursor,
        Math.max(0, capacityPt - notice.advancePt), indexOf, supports);
    if (!reserved) return null;
    if (!reserved.nextCursor) throw new Error('Footnote partition completed under a smaller capacity');
    const fragment = reserved.fragment;
    const trailing = placeNoteSeparatorOccurrence(notice, {
      occurrenceId: `${acquired.id}:notice`,
      flowDomainId: acquired.flowDomainId,
      yPt: fragment.flowBounds.yPt + fragment.advancePt,
    });
    const advancePt = fragment.advancePt + trailing.advancePt;
    const bounds = Object.freeze({ ...fragment.flowBounds, heightPt: advancePt });
    return Object.freeze({
      fragment: Object.freeze({ ...fragment, flowBounds: bounds, inkBounds: bounds, advancePt, trailing }),
      nextCursor: reserved.nextCursor,
      hostShiftPt: reserved.hostShiftPt,
    });
  };
}

function partitionIndexedFootnote(acquired: NoteLayout, cursor: FootnoteCursor | null, capacityPt: number,
  indexOf: (block: ParagraphLayout) => ParagraphContinuationIndex,
  supports: (story: StoryLayout) => boolean,
  /** Most real lines to keep; the minimum-head extent keeps exactly one. */
  lineBudget = Number.POSITIVE_INFINITY,
): FootnotePartition | null {
  if (!Number.isFinite(capacityPt) || capacityPt < 0) throw new RangeError('Invalid footnote capacity');
  // A cursor counts shaped lines, so continuing it in a different text width
  // could silently skip or repeat source text. Reflow-aware source offsets are
  // required before that case can be supported.
  if (cursor && Math.abs(cursor.inlineExtentPt - acquired.flowBounds.widthPt) > 1e-6) {
    throw new Error('Footnote continuation across different text widths is unsupported');
  }
  if (cursor) {
    const block = acquired.story.blocks[cursor.blockIndex];
    if (!block || block.kind !== 'paragraph'
      || !sameSourcePosition(cursor.sourcePosition, lineSourcePosition(block, cursor.lineIndex, indexOf(block)))) {
      throw new Error('Footnote continuation source cut changed during reflow');
    }
  }
  if (cursor === null && acquired.advancePt <= capacityPt && lineBudget === Number.POSITIVE_INFINITY) {
    return Object.freeze({ fragment: acquired, nextCursor: null, hostShiftPt: 0 });
  }
  let remainingLines = lineBudget;
  // The line cursor owns ordinary paragraph flow only. Table rows and framed
  // paragraphs need their own placement/continuation contract; a frame's
  // independent position must not be folded as an adjacent paragraph margin.
  // Relocate their reference until they fit whole, or let admission report its
  // capacity error. Never retain clipped or page-stale unsupported content.
  if (!supports(acquired.story)) return null;
  // The scalar separator band or the native leading occurrence's advance.
  const separatorPt = acquired.advancePt - acquired.story.advancePt;
  let usedPt = separatorPt;
  if (usedPt >= capacityPt) return null;
  const blocks: PaintNode[] = [];
  const startBlock = cursor?.blockIndex ?? 0;
  let nextCursor: FootnoteCursor | null = null;
  let hostShiftPt: number | null = null;
  // `projected` is `source` placed in the fragment; its first line is source
  // line `lineStart` (a line-free paragraph moves by its flow origin).
  const retain = (projected: ParagraphLayout, source: ParagraphLayout, lineStart: number) => {
    blocks.push(projected);
    if (hostShiftPt !== null
      || !projected.textBoxes.some((textBox) => textBox.story.bandDependent === true)) return;
    const first = projected.lines[0];
    hostShiftPt = first
      ? first.bounds.yPt - source.lines[lineStart]!.bounds.yPt
      : projected.flowBounds.yPt - source.flowBounds.yPt;
  };
  const cut = (block: ParagraphLayout, blockIndex: number, lineIndex: number): FootnoteCursor =>
    Object.freeze({ blockIndex, lineIndex, inlineExtentPt: acquired.flowBounds.widthPt,
      sourcePosition: lineSourcePosition(block, lineIndex, indexOf(block)) });
  for (let blockIndex = startBlock; blockIndex < acquired.story.blocks.length; blockIndex += 1) {
    const block = acquired.story.blocks[blockIndex]!;
    if (block.kind !== 'paragraph') throw new Error('Unsupported footnote block');
    // Story acquisition already folds §17.3.1.9 contextualSpacing and
    // §17.3.1.33 adjacent margins. Preserve its cursor delta within this
    // fragment rather than reintroducing each preceding space-after. A new
    // page starts a fresh band: it does not carry a prior-page overlap; before
    // and after ownership remain the shared paragraph slice contract's values.
    if (blocks.length > 0) {
      const previous = acquired.story.blocks[blockIndex - 1]!;
      usedPt += block.flowBounds.yPt - (previous.flowBounds.yPt + previous.advancePt);
    }
    const place = (paragraph: typeof block) => projectBodyOccurrence(paragraph, {
      occurrenceId: `${acquired.id}:block:${blockIndex}`,
      destination: {
        coordinateSpace: 'logical-page-points',
        flowDomainId: `${acquired.flowDomainId}:footnote:${acquired.source.storyInstance}`,
        translation: { xPt: 0,
          yPt: acquired.story.flowBounds.yPt + usedPt - separatorPt - paragraph.flowBounds.yPt },
      },
    });
    const lineStart = blockIndex === startBlock ? (cursor?.lineIndex ?? 0) : 0;
    // The minimum-head extent ends after its first real line, exactly where a
    // capacity-bound partition would cut before an unfitting next line.
    if (remainingLines <= 0) {
      nextCursor = cut(block, blockIndex, lineStart);
      break;
    }
    if (block.lines.length === 0) {
      if (usedPt + block.advancePt > capacityPt) {
        nextCursor = cut(block, blockIndex, 0);
        break;
      }
      retain(place(block), block, 0);
      usedPt += block.advancePt;
      continue;
    }
    // Select from scalar retained advances before constructing geometry.
    // Prefix slicing for every candidate duplicated all lines quadratically.
    let low = lineStart + 1;
    let high = Math.min(block.lines.length, lineStart + remainingLines);
    let admitted = lineStart;
    const index = indexOf(block);
    const first = block.lines[lineStart]!;
    const previous = block.lines[lineStart - 1];
    const leadingPt = lineStart === 0 ? block.spacing.beforePt
      + Math.max(0, first.bounds.yPt - (block.flowBounds.yPt + block.spacing.beforePt))
      : -Math.max(0, first.bounds.yPt - (previous!.bounds.yPt + previous!.advancePt));
    while (low <= high) {
      const end = Math.floor((low + high) / 2);
      const advancePt = leadingPt + index.advances[end]! - index.advances[lineStart]!
        + (end === block.lines.length ? block.spacing.afterPt : 0);
      if (usedPt + advancePt <= capacityPt + 1e-6) {
        admitted = end;
        low = end + 1;
      } else high = end - 1;
    }
    if (admitted === lineStart) {
      nextCursor = cut(block, blockIndex, lineStart);
      break;
    }
    const admittedBlock = sliceParagraphLayout(block, {
      lineStart, lineEnd: admitted,
      continuesFromPrevious: lineStart > 0, continuesOnNext: admitted < block.lines.length,
    });
    retain(place(admittedBlock), block, lineStart);
    usedPt += admittedBlock.advancePt;
    remainingLines -= admitted - lineStart;
    if (admitted < block.lines.length) {
      nextCursor = cut(block, blockIndex, admitted);
      break;
    }
  }
  if (blocks.length === 0) return null;
  const storyBounds = Object.freeze({ ...acquired.story.flowBounds, heightPt: usedPt - separatorPt });
  const story: StoryLayout = Object.freeze({
    ...acquired.story,
    blocks: Object.freeze(blocks),
    flowBounds: storyBounds,
    inkBounds: storyBounds,
    advancePt: usedPt - separatorPt,
  });
  const noteBounds = Object.freeze({ ...acquired.flowBounds, heightPt: usedPt });
  return Object.freeze({
    fragment: Object.freeze({
      ...acquired,
      story,
      flowBounds: noteBounds,
      inkBounds: noteBounds,
      advancePt: usedPt,
    }),
    nextCursor,
    hostShiftPt: hostShiftPt ?? 0,
  });
}

/**
 * ECMA-376 §17.18.33 places the continuation notice at the bottom of a page
 * on which a note continues, not after that particular note. When a later
 * note joins the band, the page's single notice moves from its carrier to
 * the new last note. Its measured advance moves with it, so the page reserve
 * that already charged it is neither repeated nor released.
 */
export function moveContinuationNotice(
  carrier: NoteLayout,
  receiver: NoteLayout,
): Readonly<{ carrier: NoteLayout; receiver: NoteLayout }> {
  const notice = carrier.trailing;
  if (!notice || receiver.trailing) throw new Error('A page note band owns exactly one continuation notice');
  const { trailing: _moved, ...retained } = carrier;
  const carrierAdvancePt = carrier.advancePt - notice.advancePt;
  const carrierBounds = Object.freeze({ ...carrier.flowBounds, heightPt: carrierAdvancePt });
  const trailing = placeNoteSeparatorOccurrence(notice, {
    occurrenceId: `${receiver.id}:notice`,
    flowDomainId: receiver.flowDomainId,
    yPt: receiver.flowBounds.yPt + receiver.advancePt,
  });
  const receiverAdvancePt = receiver.advancePt + trailing.advancePt;
  const receiverBounds = Object.freeze({ ...receiver.flowBounds, heightPt: receiverAdvancePt });
  return Object.freeze({
    carrier: Object.freeze({ ...retained, flowBounds: carrierBounds, inkBounds: carrierBounds, advancePt: carrierAdvancePt }),
    receiver: Object.freeze({
      ...receiver,
      flowBounds: receiverBounds,
      inkBounds: Object.freeze(unionLayoutRects([receiver.inkBounds, receiverBounds])!),
      advancePt: receiverAdvancePt,
      trailing,
    }),
  });
}
