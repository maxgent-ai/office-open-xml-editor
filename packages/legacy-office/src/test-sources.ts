// Test-only legacy model sources with explicit URLs: vitest resolves the
// file: source modules through its module runner and the WASM from disk.
import { createLegacyDoc } from './legacy-doc.js';
import { createLegacyPpt } from './legacy-ppt.js';
import { createLegacyXls } from './legacy-xls.js';

const here = (path: string): string => new URL(path, import.meta.url).href;

export const TEST_SOURCE_URLS = {
  doc: { wasmUrl: here('./wasm-direct-doc/legacy_doc_direct_bg.wasm'), moduleUrl: here('./legacy-doc-source-module.ts') },
  xls: { wasmUrl: here('./wasm-direct-xls/legacy_xls_direct_bg.wasm'), moduleUrl: here('./legacy-xls-source-module.ts') },
  ppt: { wasmUrl: here('./wasm-direct-ppt/legacy_ppt_direct_bg.wasm'), moduleUrl: here('./legacy-ppt-source-module.ts') },
} as const;

export const testDocSource = (options: { maxInputBytes?: number } = {}) => createLegacyDoc({ ...TEST_SOURCE_URLS.doc, ...options });
export const testXlsSource = (options: { maxInputBytes?: number } = {}) => createLegacyXls({ ...TEST_SOURCE_URLS.xls, ...options });
export const testPptSource = (options: { maxInputBytes?: number } = {}) => createLegacyPpt({ ...TEST_SOURCE_URLS.ppt, ...options });
