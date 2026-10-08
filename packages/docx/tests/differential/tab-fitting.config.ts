import { resolve } from 'node:path';
import { defineConfig } from 'vitest/config';

const base = process.env.DOCX_TAB_BASELINE_CHECKOUT;
if (!base) throw new Error('DOCX_TAB_BASELINE_CHECKOUT must name a clean origin/main checkout');

export default defineConfig({
  // Literal test imports keep the repository's module boundary gate intact.
  // Never alias these to candidate modules: this is the actual main graph.
  resolve: { alias: {
    '@docx-tab-baseline/document-layout': resolve(base, 'packages/docx/src/document-layout.ts'),
    '@docx-tab-baseline/layout-runtime': resolve(base, 'packages/docx/src/layout-runtime.ts'),
  } },
  define: { __OOXML_MODEL_SOURCES__: 'true' },
  test: { include: ['packages/docx/tests/differential/tab-fitting.test.ts'] },
});
