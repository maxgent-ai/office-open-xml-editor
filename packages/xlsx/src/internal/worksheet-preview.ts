import type { Row, Worksheet } from '../types.js';
import type { WorksheetPreviewBlocker } from '../worksheet-pull-codec.js';

type Waiter = { row: number; resolve: () => void; reject: (reason: unknown) => void };

/** One provisional model for a worksheet pull. A fallback has no provisional
 * model and makes viewers wait for the existing full-model completion path.
 * The row array is shared with the final model, but cache admission remains
 * transactional until terminal ACK. */
export class WorksheetPreview {
  readonly ready: Promise<Worksheet>;
  private resolveReady!: (worksheet: Worksheet) => void;
  private rejectReady!: (reason: unknown) => void;
  private settledReady = false;
  private waiters: Waiter[] = [];
  worksheet: Worksheet | null = null;
  reason: WorksheetPreviewBlocker | null = null;
  maxRow = 0;
  maxCol = 0;
  coveredThrough = 0;
  complete = false;
  failure: unknown;

  constructor(readonly rows: Row[]) {
    this.ready = new Promise<Worksheet>((resolve, reject) => {
      this.resolveReady = resolve;
      this.rejectReady = reject;
    });
    // Ordinary getWorksheet callers never await ready. Its rejection must not
    // become an unhandled promise when a pull fails before a viewer joins it.
    void this.ready.catch(() => undefined);
  }

  preview(worksheet: Worksheet | null, reason: WorksheetPreviewBlocker | null,
          maxRow: number, maxCol: number): void {
    this.reason = reason;
    this.maxRow = maxRow;
    this.maxCol = maxCol;
    if (!worksheet) return;
    worksheet.rows = this.rows;
    this.worksheet = worksheet;
  }

  append(rows: Row[]): void {
    if (rows.length === 0) return;
    this.rows.push(...rows);
    this.coveredThrough = rows[rows.length - 1].index;
    if (this.worksheet && !this.settledReady) this.resolveFirst(this.worksheet);
    this.flush();
  }

  finish(worksheet: Worksheet): void {
    this.complete = true;
    this.coveredThrough = Number.MAX_SAFE_INTEGER;
    if (!this.settledReady) this.resolveFirst(worksheet);
    this.flush();
  }

  fail(reason: unknown): void {
    this.failure = reason;
    if (!this.settledReady) {
      this.settledReady = true;
      this.rejectReady(reason);
    }
    for (const waiter of this.waiters.splice(0)) waiter.reject(reason);
  }

  waitFor(row: number): Promise<void> {
    if (this.failure !== undefined) return Promise.reject(this.failure);
    if (this.complete || this.coveredThrough >= row) return Promise.resolve();
    return new Promise<void>((resolve, reject) => {
      this.waiters.push({ row, resolve, reject });
    });
  }

  private resolveFirst(worksheet: Worksheet): void {
    this.settledReady = true;
    this.resolveReady(worksheet);
  }

  private flush(): void {
    const waiting = this.waiters;
    this.waiters = [];
    for (const waiter of waiting) {
      if (this.complete || this.coveredThrough >= waiter.row) waiter.resolve();
      else this.waiters.push(waiter);
    }
  }
}
