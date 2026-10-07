/** Library resource policy shared by physical pagination and full note
 * acquisition. The ceiling is independent of authored page-admission rules. */
export const MAX_BODY_LAYOUT_PAGES = 10_000;

/** Full-note acquisition needs a bound independent of point advance: zero-
 * height paragraphs can contain arbitrarily many source nodes. One million
 * source units is a library resource ceiling for this opt-in acquisition path,
 * independent of authored layout semantics. Count strings by UTF-16 length and every
 * container/primitive once per source occurrence before any shaping occurs. */
export const MAX_FOOTNOTE_ACQUISITION_UNITS = 1_000_000;

/** Operational ceiling for one complete pagination execution, including all
 * field, reserve and anchor convergence passes and every continued footnote.
 * 32 million source units permit at most 32 maximum-sized full acquisitions;
 * smaller notes share the same aggregate allowance. This is library resource
 * governance, not an Office layout rule or a wall-clock/peak-memory guarantee.
 * Charge before each whole-story cache miss; page-local field/resource notes
 * cannot multiply a permitted million-unit source by the 10,000-page ceiling.
 * Immutable whole-story cache hits cost no new full acquisition and are free. */
export const MAX_FOOTNOTE_ACQUISITION_WORK_UNITS = 32_000_000;

export interface FootnoteAcquisitionWorkBudget {
  /** Callers supply the engine's immutable story roots, never mutable inputs. */
  sourceUnits(source: object): number;
  charge(source: object): void;
}

/** Pagination-scoped, weak source-size memo plus an aggregate debit. Values do
 * not retain source roots or acquired layouts. A failed/cancelled execution
 * releases its ledger; another pagination execution receives a fresh one. */
export function createFootnoteAcquisitionWorkBudget(): FootnoteAcquisitionWorkBudget {
  const weights = new WeakMap<object, number>();
  let spent = 0;
  const sourceUnits = (source: object) => {
    let weight = weights.get(source);
    if (weight === undefined) {
      weight = assertFootnoteAcquisitionBudget(source);
      weights.set(source, weight);
    }
    return weight;
  };
  return Object.freeze({ sourceUnits, charge(source: object) {
    const weight = sourceUnits(source);
    if (weight > MAX_FOOTNOTE_ACQUISITION_WORK_UNITS - spent) {
      throw new Error(`Footnote acquisition cumulative work budget exceeded (${MAX_FOOTNOTE_ACQUISITION_WORK_UNITS} source units)`);
    }
    // Debits are never rolled back: acquisition attempted after this point may
    // fail, but its work must not become free on a convergence retry.
    spent += weight;
  } });
}

export function assertFootnoteAcquisitionBudget(source: unknown): number {
  const pending: unknown[] = [source];
  let units = 0;
  while (pending.length > 0) {
    const value = pending.pop();
    units += typeof value === 'string' ? Math.max(1, value.length) : 1;
    if (units + pending.length > MAX_FOOTNOTE_ACQUISITION_UNITS) {
      throw new Error('Footnote acquisition source budget exceeded');
    }
    if (value && typeof value === 'object') {
      // Iterate without making another array copy of a potentially large input.
      for (const key in value) {
        if (!Object.hasOwn(value, key)) continue;
        const child = (value as Record<string, unknown>)[key];
        if (++units + pending.length > MAX_FOOTNOTE_ACQUISITION_UNITS) {
          throw new Error('Footnote acquisition source budget exceeded');
        }
        pending.push(child);
      }
    }
  }
  return units;
}
