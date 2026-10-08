import type { Worksheet } from './types.js';
import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  type NormalizedXlsxWorksheetPolicy,
} from '@silurus/ooxml-core/worker';

/**
 * Realm-local association between a library-owned worksheet model and the
 * normalized worksheet policy it was admitted under.
 *
 * - Only worksheet objects created and owned by this library should be bound.
 *   Caller-supplied or caller-mutated models are never trusted to carry policy.
 * - The binding is kept in a WeakMap keyed by object identity. It is not a
 *   public field, is not serialized and does not cross realms. Structured
 *   clones, postMessage transfers, JSON round trips and any other copy produce
 *   a new identity and MUST be rebound explicitly by the owner.
 * - There is no global mutable policy. An unbound worksheet resolves to
 *   DEFAULT_XLSX_WORKSHEET_POLICY.
 * - Rebinding the same worksheet to a different policy is permitted as an
 *   internal seam; consumers that cache derived data must compare the limits
 *   they were built with.
 */
const worksheetPolicies = new WeakMap<Worksheet, NormalizedXlsxWorksheetPolicy>();

/** Bind a library-owned worksheet model to its normalized policy. */
export function bindWorksheetPolicy(
  worksheet: Worksheet,
  policy: NormalizedXlsxWorksheetPolicy,
): void {
  worksheetPolicies.set(worksheet, policy);
}

/** Resolve the bound policy, or the default policy for an unbound worksheet. */
export function getWorksheetPolicy(worksheet: Worksheet): NormalizedXlsxWorksheetPolicy {
  return worksheetPolicies.get(worksheet) ?? DEFAULT_XLSX_WORKSHEET_POLICY;
}

/**
 * Give a library-owned projection (for example a view derived from a source
 * worksheet) the same policy as its source. An unbound source leaves the
 * projection unbound so that it resolves to the default policy.
 */
export function inheritWorksheetPolicy(source: Worksheet, projection: Worksheet): void {
  const policy = worksheetPolicies.get(source);
  if (policy === undefined) {
    worksheetPolicies.delete(projection);
  } else {
    worksheetPolicies.set(projection, policy);
  }
}
