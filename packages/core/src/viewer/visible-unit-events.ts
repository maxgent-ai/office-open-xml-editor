export interface VisibleUnitRange { topIndex: number; }

/** Emits a visible-unit change only when index, count, or layout completion
 * changes. Progressive publication can grow the count while the top stays put. */
export class VisibleUnitEvents {
  private lastIndex = -1;
  private lastTotal = -1;
  private lastComplete: boolean | null = null;

  constructor(private readonly notify: (index: number, total: number, complete: boolean) => void) {}

  publish(range: VisibleUnitRange, total: number, complete: boolean): void {
    if (range.topIndex === this.lastIndex && total === this.lastTotal &&
        complete === this.lastComplete) return;
    this.lastIndex = range.topIndex;
    this.lastTotal = total;
    this.lastComplete = complete;
    this.notify(range.topIndex, total, complete);
  }

  reset(): void {
    this.lastIndex = -1;
    this.lastTotal = -1;
    this.lastComplete = null;
  }

  resetIndex(): void { this.lastIndex = -1; }
}
