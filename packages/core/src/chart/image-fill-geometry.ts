import { OoxmlDecodedImageLimitError } from '../image/pixel-budget.js';
import { MAX_CHART_IMAGE_FILL_TILES } from './resource-limits.js';

/** Library resource/IEEE-754 policy, independent of DrawingML placement rules.
 * Reject a grid whose phase cannot be represented instead of stalling on it or
 * painting it misregistered. Use a safe-integer overflow sentinel so the typed
 * error survives worker transport. */
function rejectChartImageGeometry(): never {
  throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_CHART_IMAGE_FILL_TILES, Number.MAX_SAFE_INTEGER);
}

/** ECMA-376 §20.1.8.58: alternate tiles repeat after two cells on a flipped
 * axis. Reduce the phase by that period before combining it with the destination
 * or stepping the grid; huge translated indices can neither advance nor retain
 * parity. `%` is exact, but `origin` and `step` already carry the rounding of
 * the EMU-to-pixel conversion, so the remainder can drift by about
 * |origin| x 2^-52 per factor. Paint when that drift stays below a sixteenth of
 * a pixel and an eighth of a tile; otherwise the authored phase and parity are
 * not representable, so fail closed rather than paint a misregistered grid. */
export function chartImageTilePhase(origin: number, step: number, extent: number, flip: boolean): number {
  const period = step * (flip ? 2 : 1);
  if (![origin, step, extent, period].every(Number.isFinite)
    || !(step > 0) || !(extent > 0) || extent + step === extent) rejectChartImageGeometry();
  const phaseError = (Math.abs(origin) + period) * PHASE_ROUNDING_FACTORS * Number.EPSILON;
  if (phaseError > Math.min(MAX_PHASE_ERROR_PX, step / 8)) rejectChartImageGeometry();
  return origin % period;
}

/** EMU/pt/px scaling, alignment addition and the tile period each round once. */
const PHASE_ROUNDING_FACTORS = 4;
const MAX_PHASE_ERROR_PX = 1 / 16;
