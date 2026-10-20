import type { LegacyFamily } from './legacy-source.js';

/** Only native Result failures enter here. Network/init/configuration errors
 * must bypass this decoder, even if their messages resemble a native marker.
 * UNSUPPORTED is the readers' declared representational boundary (including
 * invalid/unknown structures); it does not certify that a file is valid.
 * DOC acquisition/story projection/cursor reads, XLS metadata/host layout/sheet
 * projection/image reads, and PPT indexing/bootstrap/slide/image reads all pass
 * through DirectSourceRuntime's constructor and native-method guards.
 * Resource and encryption sites carry separate markers in Rust. Cursor misuse,
 * OUTPUT_TOO_LARGE, typed OOXML quotas and pull-credit errors remain distinct.
 *
 * v2 carries only bounded fixed metadata. Native diagnostics can include field
 * pictures and other document data, so never copy them into the signal.
 */
export function nativeRejection(error: unknown, format: LegacyFamily): unknown {
  if (typeof error !== 'string') return error;
  if (error.startsWith('LEGACY_UNSUPPORTED_ENCRYPTION:')) {
    return new Error('OOXML_SOURCE_UNSUPPORTED_ENCRYPTION');
  }
  if (error.startsWith('UNSUPPORTED:')) {
    return new Error(`MODEL_SOURCE_UNSUPPORTED:${JSON.stringify({ format, reason: 'source-subset' })}`);
  }
  return new Error(error);
}
