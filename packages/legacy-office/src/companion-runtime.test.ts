import { describe, expect, it } from 'vitest';
import { assertCompanionRuntime, CompanionRuntimeMismatchError, MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml/internal/companion';
import { COMPANION_VERSION } from '../../../src/companion-version.js';

describe('installed companion ABI', () => {
  it('accepts the matching runtime and rejects either a stale peer or protocol with a typed error', () => {
    expect(() => assertCompanionRuntime(COMPANION_VERSION, MODEL_SOURCE_MODULE_PROTOCOL)).not.toThrow();
    for (const [version, protocol] of [['0.0.0', MODEL_SOURCE_MODULE_PROTOCOL], [COMPANION_VERSION, 'future-protocol']]) {
      try {
        assertCompanionRuntime(version, protocol);
        throw new Error('mismatch was accepted');
      } catch (error) {
        expect(error).toBeInstanceOf(CompanionRuntimeMismatchError);
        expect(error).toMatchObject({ code: 'companion-runtime-mismatch' });
      }
    }
  });
});
