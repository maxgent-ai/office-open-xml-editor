import { OoxmlError } from '../errors/ooxml-error.js';
import { OoxmlDecodedImageLimitError } from '../image/pixel-budget.js';

/** Resource governance is independent of DrawingML placement. An authored
 * picture that exceeds a finite execution/allocation envelope must reject the
 * render, never masquerade as an absent fill and fall through to another paint. */
export function assertChartImageWork(work: number, limit: number): void {
  if (!Number.isSafeInteger(work) || work < 0 || work > limit) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', limit,
      Number.isSafeInteger(work) && work >= 0 ? work : Number.MAX_SAFE_INTEGER);
  }
}
export function rejectChartImageGeometry(): never {
  throw new OoxmlError('unsupported-image-effect', 'Unrepresentable chart image geometry');
}
