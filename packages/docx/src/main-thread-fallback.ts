/** Console notice for a worker-mode load that continued on the main thread. */
export function mainThreadFallbackWarning(reason: 'vertical' | 'nativeRouteSentinels'): string {
  return reason === 'nativeRouteSentinels'
    ? "[ooxml] mode: 'worker' fell back to main-thread rendering because the worker cannot register the font probes that size absent-font lines."
    : "[ooxml] mode: 'worker' fell back to main-thread rendering because this document requires DOM OpenType vertical glyph selection.";
}
