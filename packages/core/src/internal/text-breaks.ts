// Pure wrapping primitives for format adapters, including worker/Node renderers.
// Keep this entry independent of the root rendering barrel.
export { isCjkBreakChar, isLatinWordCodePoint } from '../text/cjk-ranges.js';
export { isUax14NoBreakPair } from '../text/line-break.js';
export { kinsokuAdjustedSplit, DEFAULT_KINSOKU_RULES } from '../text/kinsoku/index.js';
export {
  containsSeaScript, isGraphemeFillText, seaMixedBreakOffsets,
  fitSeaWordPrefix, graphemeClusterOffsets,
} from '../text/sea-break.js';
