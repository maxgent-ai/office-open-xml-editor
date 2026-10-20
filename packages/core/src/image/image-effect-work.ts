import { MAX_IMAGE_EFFECT_PIXEL_WORK, OoxmlDecodedImageLimitError } from './pixel-budget.js';

/** Execution ledger, charged before each kernel/copy/draw, using the actual
 * buffer grid. Planning reserves work early; this second admission boundary
 * fails closed even if a future planner transition underestimates execution.
 * Nested decoded fills share a ledger in addition to each image's plan. */
export class ImageEffectWorkLedger {
  used = 0;
  constructor(readonly limit = MAX_IMAGE_EFFECT_PIXEL_WORK) {}

  charge(work: number): void {
    const next = this.used + work;
    if (!Number.isSafeInteger(work) || work < 0 || !Number.isSafeInteger(next) || next > this.limit) {
      throw new OoxmlDecodedImageLimitError('image-effect-work', this.limit,
        Number.isSafeInteger(next) ? next : Number.MAX_SAFE_INTEGER);
    }
    this.used = next;
  }
}
