/**
 * Internal ABI for exact-version companion packages, not an application API.
 * The installed viewer owns protocol selection and all worker transport. A
 * companion bundles only stateless URL/CFB helpers and realm-local WASM guards:
 * guards classify traps by error name/code across self-contained module realms,
 * never by a shared class identity. ModelSource types are structural and can
 * be bundled in declarations. No bridge or transport is bundled there.
 */
import { MODEL_SOURCE_MODULE_PROTOCOL } from '../packages/core/src/source/model-source.js';
import { COMPANION_VERSION } from './companion-version.js';
export { MODEL_SOURCE_MODULE_PROTOCOL };

export class CompanionRuntimeMismatchError extends Error {
  readonly code = 'companion-runtime-mismatch';
  constructor(expectedVersion: string, expectedProtocol: string) {
    super(`Companion requires @silurus/ooxml ${expectedVersion} (${expectedProtocol}); installed runtime is ${COMPANION_VERSION} (${MODEL_SOURCE_MODULE_PROTOCOL})`);
    this.name = 'CompanionRuntimeMismatchError';
  }
}

/** Fail before claiming input or sending a descriptor to a worker. */
export function assertCompanionRuntime(version: string, protocol: string): void {
  if (version !== COMPANION_VERSION || protocol !== MODEL_SOURCE_MODULE_PROTOCOL) {
    throw new CompanionRuntimeMismatchError(version, protocol);
  }
}
