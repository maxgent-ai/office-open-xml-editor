// Private session metadata, absent from the public worksheet and load options.
// Each model owner binds the load-time value after decoding; projections inherit
// it explicitly. The render worker receives the caller's value at parse time
// and never reads its own navigator for this choice.
const japaneseCellFonts = new WeakMap<object, boolean>();

export function snapshotJapaneseCellFonts(): boolean {
  return typeof navigator !== 'undefined'
    && (navigator.language || '').toLowerCase().startsWith('ja');
}

export function bindCellFontLanguage(model: object, japanese: boolean): void {
  japaneseCellFonts.set(model, japanese === true);
}

/** Loaded workbooks bind this before exposing any sheet, including previews
 * and sheets loaded later. Standalone renderer models snapshot at first use.
 * Re-open the workbook to apply a change to the viewer language; layout and
 * paint retain one choice for the entire opened workbook's lifetime. */
export function usesJapaneseCellFonts(model: object | undefined): boolean {
  if (!model) return false;
  let japanese = japaneseCellFonts.get(model);
  if (japanese === undefined) {
    japanese = snapshotJapaneseCellFonts();
    japaneseCellFonts.set(model, japanese);
  }
  return japanese;
}

export function inheritCellFontLanguage(source: object, projection: object): void {
  bindCellFontLanguage(projection, usesJapaneseCellFonts(source));
}
