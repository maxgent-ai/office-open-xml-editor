import { vi } from 'vitest';
import { documentLayoutValidationEnabled } from '../layout/validation-policy.js';

/** Test-only control of the redundant diagnostic walk. */
export function setDocumentLayoutValidation(enabled: boolean): void {
  vi.mocked(documentLayoutValidationEnabled).mockReturnValue(enabled);
}
