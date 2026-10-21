import { describe, it, expect } from 'vitest';
import {
  kashidaInsertionPoints,
  priorityKashidaInsertionCandidates,
  priorityKashidaInsertionPoints,
} from './arabic-joining.js';

// Arabic letters by Unicode Joining_Type (UCD DerivedJoiningType.txt), spelled
// with \u escapes so the source stays visually LTR.
const BEH = 'ب'; // D  Dual_Joining
const LAM = 'ل'; // D  Dual_Joining
const SEEN = 'س'; // D Dual_Joining
const ALEF = 'ا'; // R  Right_Joining
const REH = 'ر'; // R  Right_Joining
const WAW = 'و'; // R  Right_Joining
const FATHA = 'ً'; // T Transparent (harakat)
const TATWEEL = 'ـ'; // C Join_Causing
const ALEF_MADDA = 'آ'; // R (alef with madda above) — a lam-alef partner

// kashidaInsertionPoints returns code-point offsets `k` (1 ≤ k ≤ len-1) meaning
// "a tatweel may be inserted BEFORE the code point at index k" (equivalently,
// after the letter cluster ending at k-1). This matches the `splitBefore`
// offset convention of the shared line-distribute kernel.
describe('kashidaInsertionPoints', () => {
  it('opens a point between two dual-joining letters', () => {
    expect(kashidaInsertionPoints(BEH + BEH)).toEqual([1]);
    expect(kashidaInsertionPoints(BEH + BEH + BEH)).toEqual([1, 2]);
    expect(kashidaInsertionPoints(SEEN + BEH + LAM + BEH)).toEqual([1, 2, 3]);
  });
  it('opens a point before a right-joining letter that follows a joiner (beh-alef)', () => {
    // BEH joins-to-following, ALEF joins-to-preceding → the boundary connects.
    expect(kashidaInsertionPoints(BEH + ALEF)).toEqual([1]);
  });
  it('opens NO point after a right-joining letter (it does not join forward)', () => {
    // ALEF (R) does not join to the following BEH → no point at offset 1.
    expect(kashidaInsertionPoints(ALEF + BEH)).toEqual([]);
    // …but a subsequent BEH-BEH boundary is still eligible.
    expect(kashidaInsertionPoints(ALEF + BEH + BEH)).toEqual([2]);
    // REH / WAW behave like ALEF (right-joining).
    expect(kashidaInsertionPoints(REH + BEH + BEH)).toEqual([2]);
    expect(kashidaInsertionPoints(BEH + WAW + BEH)).toEqual([1]);
  });
  it('EXCLUDES the lam-alef ligature boundary', () => {
    // LAM (D) + ALEF (R) would be eligible by the generic rule, but lam-alef is a
    // mandatory ligature — Word does not insert kashida inside it.
    expect(kashidaInsertionPoints(LAM + ALEF)).toEqual([]);
    expect(kashidaInsertionPoints(LAM + ALEF_MADDA)).toEqual([]);
    // A lam that is NOT before an alef still opens a point.
    expect(kashidaInsertionPoints(LAM + BEH)).toEqual([1]);
    // beh-lam-alef: the beh-lam boundary (offset 1) is eligible; lam-alef (offset 2) is not.
    expect(kashidaInsertionPoints(BEH + LAM + ALEF)).toEqual([1]);
  });
  it('skips Transparent marks (harakat) for adjacency', () => {
    // BEH + FATHA + BEH: the mark belongs to the first BEH; the join is BEH→BEH
    // across it, so the only point is BEFORE the second BEH (offset 2). No point
    // is opened before the mark itself (offset 1).
    expect(kashidaInsertionPoints(BEH + FATHA + BEH)).toEqual([2]);
  });
  it('returns nothing for non-Arabic text', () => {
    expect(kashidaInsertionPoints('abcd')).toEqual([]);
    expect(kashidaInsertionPoints('')).toEqual([]);
    expect(kashidaInsertionPoints(BEH)).toEqual([]);
  });
});

describe('priorityKashidaInsertionPoints', () => {
  it('chooses one highest-priority join per whitespace-delimited word', () => {
    expect(priorityKashidaInsertionPoints('السلسلة')).toEqual([5]); // Seen, nearest end
    expect(priorityKashidaInsertionPoints('بين')).toEqual([1]); // BaRa over final Normal
    expect(priorityKashidaInsertionPoints('الحروف')).toEqual([3]); // HahDal
    expect(priorityKashidaInsertionPoints(`${BEH}${BEH} ${BEH}${BEH}`)).toEqual([1, 4]);
  });

  it('prefers a join after a source tatweel', () => {
    expect(priorityKashidaInsertionPoints(BEH + TATWEEL + BEH + BEH)).toEqual([2]);
  });

  it('maps final Tah and Qaf aliases to their documented priority classes', () => {
    expect(priorityKashidaInsertionPoints('بيط')).toEqual([2]);
    expect(priorityKashidaInsertionCandidates('بيط')).toEqual([
      { beforeCp: 2, priority: 10 },
    ]);
    expect(priorityKashidaInsertionCandidates(BEH + BEH + 'ق')).toEqual([
      { beforeCp: 2, priority: 8 },
    ]);
  });

  it('does not classify non-BEH joining groups as BaRa predecessors', () => {
    // NOON->YEH is eligible but both joins are Normal, so the word-end tie-break wins.
    expect(priorityKashidaInsertionPoints('نيب')).toEqual([2]);
  });
});
