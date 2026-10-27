/**
 * Admission of one `fontResolver` answer, shared by the realm that paints
 * (`source-session.ts`) and by the main-thread side of the worker bridge
 * (`resolver-bridge-responder.ts`), so a resolver answer is judged by the same
 * rules whether the function runs in the painting realm or is bridged to it.
 * Library resource policy, not OOXML behaviour.
 */
import { fontSourceLimitError } from '../errors/font-source-limit.js';
import type { OoxmlFontSourceViolation } from '../errors/ooxml-error.js';
import { HARD_MAX_FONT_SOURCE_DECLARATIONS } from '../worker/resource-policy.generated.js';
import { validateFontFaceDeclaration, type FontFaceDeclaration } from './source.js';

/** A resolver answer: validated declarations, or a failure reason that keeps
 * every demanded family on its host fallback. */
export type AdmittedFontResolverResult =
  | { readonly ok: true; readonly declarations: readonly FontFaceDeclaration[] }
  | { readonly ok: false; readonly reason: string };

/**
 * Validate what a resolver function returned. A non-array or a malformed
 * declaration is a failure (`ok: false`, the families keep their host
 * fallback); more than `HARD_MAX_FONT_SOURCE_DECLARATIONS` entries or a `data`
 * buffer beyond the face ceiling throws `OoxmlError('font-source-limit')`,
 * which rejects the load. Declarations are validated in order, so the first
 * offending entry decides between the two.
 */
export function admitFontResolverResult(raw: unknown): AdmittedFontResolverResult {
  if (!Array.isArray(raw)) return { ok: false, reason: 'resolver did not return an array' };
  if (raw.length > HARD_MAX_FONT_SOURCE_DECLARATIONS) {
    throw fontSourceLimitError('declarations', HARD_MAX_FONT_SOURCE_DECLARATIONS, raw.length);
  }
  try {
    return {
      ok: true,
      declarations: Object.freeze(raw.map((entry: unknown, index) =>
        validateFontFaceDeclaration(entry, `font resolver declaration ${index}`))),
    };
  } catch (error) {
    if (!(error instanceof TypeError)) throw error;
    return { ok: false, reason: error.message };
  }
}

/** Limit errors a bridged resolver call reports for the main-side admission:
 * the painting realm rethrows them as its own admission would have. */
const bridgedLimitErrors = new WeakSet<object>();

/** The `font-source-limit` rejection the main-side admission of a bridged
 * resolver call raised, re-created in the painting realm. */
export function bridgedFontSourceLimitError(violation: OoxmlFontSourceViolation): Error {
  const error = fontSourceLimitError(violation.quota, violation.limit, violation.observed);
  bridgedLimitErrors.add(error);
  return error;
}

/** Whether `error` is a {@link bridgedFontSourceLimitError}. An application
 * resolver that throws an `OoxmlError` itself only fails its families. */
export function isBridgedFontSourceLimitError(error: unknown): boolean {
  return typeof error === 'object' && error !== null && bridgedLimitErrors.has(error);
}
