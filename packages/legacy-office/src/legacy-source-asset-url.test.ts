import { afterEach, describe, expect, it, vi } from 'vitest';
// Public synthetic bundler output: no native reader or binary is imported.
vi.mock('./wasm-direct-doc/legacy_doc_direct_bg.wasm?url', () => ({ default: '/_next/static/media/doc.wasm' }));
vi.mock('./legacy-doc-source-module.ts?url', () => ({ default: '/_next/static/media/doc.mjs' }));
vi.mock('./wasm-direct-xls/legacy_xls_direct_bg.wasm?url', () => ({ default: '/_next/static/media/xls.wasm' }));
vi.mock('./legacy-xls-source-module.ts?url', () => ({ default: '/_next/static/media/xls.mjs' }));
vi.mock('./wasm-direct-ppt/legacy_ppt_direct_bg.wasm?url', () => ({ default: '/_next/static/media/ppt.wasm' }));
vi.mock('./legacy-ppt-source-module.ts?url', () => ({ default: '/_next/static/media/ppt.mjs' }));

import { createLegacyDoc, legacyDoc } from './legacy-doc.js';
import { createLegacyXls, legacyXls } from './legacy-xls.js';
import { createLegacyPpt, legacyPpt } from './legacy-ppt.js';

afterEach(() => vi.unstubAllGlobals());

// The calling realm owns browser asset URL resolution, even when the bundler
// retains a build-time file URL as import.meta.url. No source/WASM is loaded.
describe.each([
  { family: 'doc', create: createLegacyDoc, constant: legacyDoc },
  { family: 'xls', create: createLegacyXls, constant: legacyXls },
  { family: 'ppt', create: createLegacyPpt, constant: legacyPpt },
])('legacy $family emitted asset URLs', ({ family, create, constant }) => {
  it('describes browser-fetchable default module and WASM URLs before worker handoff', () => {
    vi.stubGlobal('location', { href: 'https://app.example.test/reports/current' });
    const { module } = create().beginLoad();
    expect(module.moduleUrl).toBe(`https://app.example.test/_next/static/media/${family}.mjs`);
    expect(module.config.wasmUrl).toBe(`https://app.example.test/_next/static/media/${family}.wasm`);
  });

  // The constant was imported before `location` existed: it resolves the
  // bundled URLs on first use, against the page that uses it.
  it('resolves the default constant\'s asset URLs on first use, not at import', () => {
    vi.stubGlobal('location', { href: 'https://app.example.test/reports/current' });
    const { module } = constant.beginLoad();
    expect(module.moduleUrl).toBe(`https://app.example.test/_next/static/media/${family}.mjs`);
    expect(module.config.wasmUrl).toBe(`https://app.example.test/_next/static/media/${family}.wasm`);
  });
});
