import type { TextRunData } from '@silurus/ooxml-core';

/** Shared display preparation. DrawingML caps and slidenum field substitution
 * keep the renderer's existing non-locale uppercase and numbering behavior. */
export function preparedPowerPointText(run: TextRunData, slideNumber?: number): string {
  const text = run.fieldType === 'slidenum' && slideNumber !== undefined ? String(slideNumber) : run.text;
  return run.caps === 'all' || run.caps === 'small' ? text.toUpperCase() : text;
}
