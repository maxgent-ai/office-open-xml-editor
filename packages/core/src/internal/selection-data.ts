/**
 * Selection-context data attributes on text-layer elements. Viewers write the
 * run's unit/run indices as `data-*` attributes and read them back to resolve
 * a DOM selection to model coordinates.
 */

/** Set (or, with `value` undefined, remove) one selection data attribute. A
 *  minimal element without `dataset` falls back to `setAttribute`. */
export function setSelectionData(element: HTMLElement, key: string, value?: string): void {
  if (element.dataset) {
    if (value === undefined) delete element.dataset[key];
    else element.dataset[key] = value;
  } else if (value !== undefined) {
    element.setAttribute?.(`data-${key.replace(/[A-Z]/g, (char) => `-${char.toLowerCase()}`)}`, value);
  }
}

/** Parse a decimal data attribute as a non-negative safe integer, or `null`. */
export function nonNegativeInteger(value: string | undefined): number | null {
  if (value === undefined || !/^\d+$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) ? number : null;
}
