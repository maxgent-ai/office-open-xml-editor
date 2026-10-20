/**
 * Note-area arithmetic, independent of note/separator and page-placement policy.
 * Reduce advances from zero in source order, then add the area origin once per
 * boundary. Accumulating a page-coordinate cursor rounds each addition at page
 * magnitude and can drift beyond an area whose height is reduced from zero.
 *
 * The terminal prefix is both the area's advance and its final bottom offset.
 * Acquisition appends advances as they become known; composition can rebase the
 * same prefixes without changing their reduction order. This is numerical
 * consistency policy; it adds no layout rule and no fit tolerance.
 */
export class NoteStack {
  private readonly prefixAdvancesPt: number[] = [0];

  get advancePt(): number {
    return this.prefixAdvancesPt[this.prefixAdvancesPt.length - 1]!;
  }

  append(advancePt: number): void {
    this.prefixAdvancesPt.push(this.advancePt + advancePt);
  }

  /** Index addresses a note's top, or the terminal bottom at the note count. */
  topAt(areaTopPt: number, index = this.prefixAdvancesPt.length - 1): number {
    return areaTopPt + this.prefixAdvancesPt[index]!;
  }
}
