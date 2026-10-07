import { inverseMapAxisAlignedRect } from './affine.js';
import { ExactConvergenceError } from './convergence.js';
import type { LayoutRect, Matrix2DData } from './types.js';

/*
 * Destination page frames of a story laid out before its composer places it.
 *
 * Content that only a page position places — §17.4.57 positioned tables of a
 * story table (table-pagination.ts layoutWholeTableOnPage) and the
 * page/margin axes of a header or footer story-root cell-owner host
 * (table-owner-runs.ts) — is resolved against the page while its story is
 * laid out. A header or footer is placed by a band translation that depends
 * on its own extent, and a text box places its story by the story's extent
 * (autofit, anchor offset), so both are exact fixed points solved here. These
 * are library choices composing the existing transforms; no Office control
 * establishes them.
 */

/** Page and margin rectangles of a destination page, page origin at 0. */
export interface StoryPageFrames {
  readonly page: LayoutRect;
  readonly margin: LayoutRect;
  /**
   * A text box story's frames only: the translation (story coordinates) its
   * box gives the laid-out story's flow, the anchor offset and fitted shift.
   * Page-owned axes of a drawing anchored in the story do not receive it
   * (retained-geometry-translation.ts translateParagraphLayout), so the
   * frames of those axes are these moved back by it ({@link storyAnchorPageFrames}).
   */
  readonly storyFlowPt?: Readonly<{ xPt: number; yPt: number }>;
}

/**
 * The page frames in the coordinates a text box story's page-owned anchor
 * axes keep, and the translation its host-following content still receives
 * to reach them: the pair a drawing anchored in that story resolves its own
 * text box story's page frames through (ParagraphAcquisitionOptions.hostPageFrames
 * and hostFlowPageTranslationPt), exactly as a body or page-story paragraph's
 * drawing does with the page and its band.
 */
export function storyAnchorPageFrames(frames: StoryPageFrames): Readonly<{
  frames: StoryPageFrames;
  flowPt: Readonly<{ xPt: number; yPt: number }>;
}> {
  const flowPt = frames.storyFlowPt ?? { xPt: 0, yPt: 0 };
  const move = (rect: LayoutRect) => Object.freeze({
    ...rect, xPt: rect.xPt + flowPt.xPt, yPt: rect.yPt + flowPt.yPt,
  });
  return Object.freeze({
    frames: Object.freeze({ page: move(frames.page), margin: move(frames.margin) }),
    flowPt,
  });
}

/**
 * The destination page frames stated in a frame whose coordinates
 * `toFrames` carries to the frames' own (a text box story's placement on the
 * page): each rectangle carried back through it, so content resolved against
 * them in the story's axes is the page's whatever orientation the story has.
 * Null when the placement is not a translation, scale or exact quarter turn:
 * a page band carried into a story turned by any other angle is a slanted
 * strip, not a band. Such a story is given no page frames, so its
 * page-placed content keeps the placement it has without them (positioned
 * tables of its tables are not given final frames, as before frames were
 * carried into stories at all); no bounding-box approximation is adopted.
 */
export function storyPageFramesThrough(
  frames: StoryPageFrames,
  toFrames: Matrix2DData,
): StoryPageFrames | null {
  const page = inverseMapAxisAlignedRect(toFrames, frames.page);
  const margin = inverseMapAxisAlignedRect(toFrames, frames.margin);
  return page && margin ? Object.freeze({ page, margin }) : null;
}

/**
 * The exact fixed point of a placement that carries the destination page
 * frames into a story's coordinates and itself depends on that story (a text
 * box positions its story by the story's extent, which depends on the
 * page-placed content laid out against those frames). `evaluate` lays the
 * story out against `frames` and states the frames its resulting placement
 * implies; the value returned is that of an evaluation whose implied frames
 * equal, exactly, the frames it was laid out with. Nothing else is accepted.
 *
 * Only the search for that point differs from plain repetition. A box moves
 * its story along one axis, so the implied frames are the start frames
 * translated along that axis by a scalar t, and the fixed point is a root of
 * the residual r(t) = t' − t, searched by {@link translationRootSearch}.
 * Frames not of that one-axis family are repeated plainly. `limit` counts
 * evaluations and guards resources only: exhaustion, or a repeated candidate
 * under plain repetition, fails closed with {@link ExactConvergenceError}, as
 * a residual whose sign change is a jump (no exact fixed point) does.
 */
export function solveStoryPageFrames<T>(
  start: StoryPageFrames,
  evaluate: (frames: StoryPageFrames | null) => Readonly<{ value: T; next: StoryPageFrames | null }>,
  limit: number,
): T {
  const key = (frames: StoryPageFrames | null) => JSON.stringify(frames);
  let axis: 'xPt' | 'yPt' | null = null;
  // The offset of `frames` from `start` along the one axis, or null when
  // `frames` is not such a translation.
  const offsetOf = (frames: StoryPageFrames | null): number | null => {
    if (!frames) return null;
    const dx = frames.page.xPt - start.page.xPt;
    const dy = frames.page.yPt - start.page.yPt;
    const translated = (from: LayoutRect, to: LayoutRect) => to.widthPt === from.widthPt
      && to.heightPt === from.heightPt && to.xPt === from.xPt + dx && to.yPt === from.yPt + dy;
    if (!translated(start.page, frames.page) || !translated(start.margin, frames.margin)) return null;
    // A story flow moved by −(dx, dy): the frames of the story's page-owned
    // anchor axes are unchanged (storyAnchorPageFrames).
    if ((start.storyFlowPt === undefined) !== (frames.storyFlowPt === undefined)) return null;
    if (start.storyFlowPt && frames.storyFlowPt && (
      frames.storyFlowPt.xPt !== start.storyFlowPt.xPt - dx
      || frames.storyFlowPt.yPt !== start.storyFlowPt.yPt - dy
    )) return null;
    if (dx !== 0 && dy !== 0) return null;
    if (dx === 0 && dy === 0) return 0;
    const moved = dx !== 0 ? 'xPt' : 'yPt';
    if (axis !== null && axis !== moved) return null;
    axis = moved;
    return moved === 'xPt' ? dx : dy;
  };
  const at = (t: number): StoryPageFrames => {
    const dx = axis === 'xPt' ? t : 0;
    const dy = axis === 'yPt' ? t : 0;
    const move = (rect: LayoutRect) => Object.freeze({ ...rect, xPt: rect.xPt + dx, yPt: rect.yPt + dy });
    return Object.freeze({
      page: move(start.page),
      margin: move(start.margin),
      ...(start.storyFlowPt ? {
        storyFlowPt: Object.freeze({ xPt: start.storyFlowPt.xPt - dx, yPt: start.storyFlowPt.yPt - dy }),
      } : {}),
    });
  };
  const states: string[] = [];
  const seen = new Set<string>();
  const fresh = (candidate: number) => Number.isFinite(candidate) && !seen.has(key(at(candidate)));
  const search = translationRootSearch(fresh);
  let frames: StoryPageFrames | null = start;
  for (let pass = 1; pass <= limit; pass += 1) {
    const state = key(frames);
    states.push(state);
    seen.add(state);
    const { value, next } = evaluate(frames);
    if (key(next) === state) return value;
    const t = offsetOf(frames);
    const implied = offsetOf(next);
    if (t === null || implied === null || axis === null) {
      // Plain repetition outside the one-axis family.
      if (seen.has(key(next))) throw new ExactConvergenceError('cycle', [...states, key(next)], pass);
      frames = next;
      continue;
    }
    const candidate = search(t, implied);
    if (!fresh(candidate)) {
      throw new ExactConvergenceError('cycle', [...states, key(at(candidate))], pass);
    }
    frames = at(candidate);
  }
  throw new ExactConvergenceError('limit', states, limit);
}

/**
 * The exact fixed point t = implied(t) of a placement that is a translation
 * along one axis by a scalar t and itself depends on the content it places
 * (a footer band: its translation is the page bottom less footerDistance less
 * the story extent, which page-placed content of the story — a story-root
 * host's wrap exclusion, a positioned table's anchor wrap — makes depend on
 * that translation). The search is that of {@link solveStoryPageFrames} on
 * one scalar: every candidate is an ordinary evaluation and only an
 * evaluation whose implied translation equals, exactly, the one it was laid
 * out with is accepted. `limit` counts evaluations and guards resources only;
 * exhaustion, a repeated candidate, or a residual whose sign change is a jump
 * (no exact fixed point) fails closed with {@link ExactConvergenceError}.
 */
export function solveExactTranslation<T>(
  start: number,
  evaluate: (t: number) => Readonly<{ value: T; implied: number }>,
  limit: number,
): T {
  const states: string[] = [];
  const seen = new Set<number>();
  const fresh = (candidate: number) => Number.isFinite(candidate) && !seen.has(candidate);
  const search = translationRootSearch(fresh);
  let t = start;
  for (let pass = 1; pass <= limit; pass += 1) {
    states.push(JSON.stringify(t));
    seen.add(t);
    const { value, implied } = evaluate(t);
    if (implied === t) return value;
    const candidate = search(t, implied);
    if (!fresh(candidate)) {
      throw new ExactConvergenceError('cycle', [...states, JSON.stringify(candidate)], pass);
    }
    t = candidate;
  }
  throw new ExactConvergenceError('limit', states, limit);
}

/**
 * The candidate search shared by the one-axis solvers: given each evaluated
 * t and its implied translation, the next t to evaluate. Where the placed
 * content follows t one-for-one over a range, plain repetition only moves t
 * by a constant step per pass; the search widens the step until the residual
 * r = implied − t changes sign, then narrows the bracket by the secant
 * through it (exact on one linear piece) and by the implied translation of
 * its newest sample (exact where the content no longer depends on t). Every
 * candidate is still an ordinary evaluation accepted only when exact.
 * `fresh` rejects candidates already evaluated; the caller fails closed on a
 * candidate that is not fresh.
 */
function translationRootSearch(fresh: (candidate: number) => boolean): (t: number, implied: number) => number {
  const samples: { t: number; r: number }[] = [];
  let picardLast = false;
  let widening = 1;
  return (t, implied) => {
    samples.push({ t, r: implied - t });
    // The tightest sign change among the samples (at most `limit` of them).
    let positive: { t: number; r: number } | null = null;
    let negative: { t: number; r: number } | null = null;
    for (const above of samples) {
      if (above.r <= 0) continue;
      for (const below of samples) {
        if (below.r >= 0) continue;
        if (!positive || !negative
          || Math.abs(above.t - below.t) < Math.abs(positive.t - negative.t)) {
          positive = above;
          negative = below;
        }
      }
    }
    if (positive && negative) {
      const low = Math.min(positive.t, negative.t);
      const high = Math.max(positive.t, negative.t);
      const inside = (value: number) => value > low && value < high && fresh(value);
      const secant = positive.t - positive.r * (negative.t - positive.t) / (negative.r - positive.r);
      if (!picardLast && inside(implied)) {
        picardLast = true;
        return implied;
      }
      picardLast = false;
      return inside(secant) ? secant : (low + high) / 2;
    }
    if (samples.length === 1) return implied;
    // No sign change yet: extrapolate the secant of the last two samples
    // when it points the way the residual does, else widen the step.
    const [previous, last] = samples.slice(-2) as [{ t: number; r: number }, { t: number; r: number }];
    const secant = last.r !== previous.r
      ? last.t - last.r * (last.t - previous.t) / (last.r - previous.r)
      : Number.NaN;
    widening *= 2;
    return Number.isFinite(secant) && Math.sign(secant - last.t) === Math.sign(last.r)
      ? secant
      : last.t + widening * last.r;
  };
}
