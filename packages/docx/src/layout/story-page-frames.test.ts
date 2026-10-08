import { describe, expect, it } from 'vitest';
import { ExactConvergenceError } from './convergence.js';
import { solveExactTranslation, solveStoryPageFrames, type StoryPageFrames } from './story-page-frames.js';

// Synthetic placements: the implied frames are the laid-out frames moved
// along y to `implied(t)`, t being their y offset from the start frames.
// Only exact evaluations are ever accepted; these cases pin the search that
// reaches them and the fail-closed exits, not any layout policy.
describe('story page-frame solvers', () => {
  const start: StoryPageFrames = {
    page: { xPt: 0, yPt: 0, widthPt: 400, heightPt: 400 },
    margin: { xPt: 40, yPt: 40, widthPt: 320, heightPt: 320 },
  };
  const moved = (frames: StoryPageFrames, t: number): StoryPageFrames => ({
    page: { ...frames.page, yPt: start.page.yPt + t },
    margin: { ...frames.margin, yPt: start.margin.yPt + t },
  });
  const solve = (implied: (t: number) => number) => {
    const evaluated: number[] = [];
    const value = solveStoryPageFrames(start, (frames) => {
      const t = frames!.page.yPt - start.page.yPt;
      evaluated.push(t);
      return { value: t, next: moved(frames!, implied(t)) };
    }, 16);
    return { value, evaluated };
  };

  it('reaches a fixed point past a constant-step creep and on a contracting line, exactly', () => {
    // One 8pt step per pass until the content no longer depends on t: plain
    // repetition needs 16 passes and a confirming one.
    const creep = solve((t) => Math.max(t - 8, -127));
    expect(creep.value).toBe(-127);
    expect(creep.evaluated.length).toBeLessThan(16);
    // A center anchor halves the distance per pass.
    expect(solve((t) => (t - 100) / 2).value).toBe(-100);
  });

  it('fails closed when the residual changes sign only by a jump', () => {
    expect(() => solve((t) => (t < -50 ? t + 5 : t - 5))).toThrow(ExactConvergenceError);
  });

  it('solves a scalar band translation with the same search and fails closed on a jump', () => {
    const scalar = (implied: (t: number) => number) => {
      const evaluated: number[] = [];
      const value = solveExactTranslation(20, (t) => {
        evaluated.push(t);
        return { value: t, implied: implied(t) };
      }, 16);
      return { value, evaluated };
    };
    // Constant +20 up to t = 309, then fixed at 331: plain repetition needs
    // sixteen passes before reaching it.
    const creep = scalar((t) => (t < 309 ? t + 20 : 331));
    expect(creep.value).toBe(331);
    expect(creep.evaluated.length).toBeLessThan(16);
    // +20 / −4 pieces meeting by a jump have no exact fixed point.
    expect(() => scalar((t) => (t <= 242 ? t + 20 : t - 4))).toThrow(ExactConvergenceError);
  });

  it('repeats plainly when the implied frames are not a translation along one axis', () => {
    // Sizes change: no scalar family, a fixed point still accepted exactly.
    const sized = solveStoryPageFrames(start, (frames) => ({
      value: frames!.page.heightPt,
      next: { ...frames!, page: { ...frames!.page, heightPt: Math.max(frames!.page.heightPt - 100, 200) } },
    }), 16);
    expect(sized).toBe(200);
  });
});
