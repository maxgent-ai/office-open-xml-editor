import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';

function source(relative: string): string {
  return readFileSync(new URL(relative, import.meta.url), 'utf8');
}

describe('Node facade ownership boundary', () => {
  for (const format of ['docx', 'pptx', 'xlsx'] as const) {
    it(`${format} delegates archive/runtime acquisition to the format package`, () => {
      const facade = source(`./${format}.ts`);
      expect(facade).toContain(`@silurus/ooxml-${format}/internal/session`);
      expect(facade).not.toMatch(new RegExp(`@silurus/ooxml-${format}/wasm(?:['"]|$)`));
      expect(facade).not.toContain('resourcePolicyForWasm');
      expect(facade).not.toContain('normalizeLoadResourceOptions');
      expect(facade).not.toMatch(/new\s+\w*Archive\s*\(/);

      const acquisition = source(`../../${format}/src/internal/node-acquisition.ts`);
      // The format package owns its realm runtime host (core's singleton factory).
      expect(acquisition).toContain('wasmRuntimeSingleton');
      expect(acquisition).toContain('OoxmlResourceMetricsSession');
    });
  }
});
