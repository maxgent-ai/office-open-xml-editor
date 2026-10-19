import { OoxmlError, OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import { isOoxmlDecodedImageLimitError } from './pixel-budget.js';
import { isTiffDecodeError } from './tiff-contract.js';
import { isOptionalImageCodecUnavailableError } from './optional-image-fallback.js';
import { isOoxmlIncompleteMetafileError } from './raster-or-metafile.js';

/** Strict image chains must reach every host/worker fallback as a terminal
 * diagnostic. Preserve existing resource/codec identities; native decode,
 * dependency, Canvas and bitmap failures must never become a skipped picture. */
export function imageEffectError(error: unknown): unknown {
  if (error instanceof OoxmlError || error instanceof OoxmlResourceLimitError
    || isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)
    || isOptionalImageCodecUnavailableError(error) || isOoxmlIncompleteMetafileError(error)) return error;
  return new OoxmlError('unsupported-image-effect', 'Image effect preparation or evaluation failed');
}
