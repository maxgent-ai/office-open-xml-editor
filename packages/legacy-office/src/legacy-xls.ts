import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, deferredLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-xls/legacy_xls_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-xls-source-module.ts?url';

export type LegacyXlsOptions = LegacySourceOptions;

/**
 * Create a model source for legacy Excel 97-2003 (.xls) input with custom
 * options, for example `wasmUrl` / `moduleUrl` when an asset pipeline
 * relocates the bundled files. Pass it as the `legacyXls` option of the XLSX
 * loaders and viewers. Creating it fetches nothing; the source module and its
 * WASM load only when a .xls input is claimed.
 */
export function createLegacyXls(options: LegacyXlsOptions = {}): ModelSource<'xlsx'> {
  return createLegacySource<'xlsx'>('xls', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * Legacy Excel 97-2003 (.xls) reader with the bundled asset URLs:
 * `import { legacyXls } from '@silurus/ooxml-legacy-office/xls'`, then pass
 * `{ legacyXls }` to the XLSX loaders and viewers. Importing it fetches and
 * instantiates nothing; the bundled URLs resolve on first use, and the source
 * module and its WASM load only when a .xls input is claimed. Use
 * {@link createLegacyXls} for custom asset URLs or an input size limit.
 */
export const legacyXls: ModelSource<'xlsx'> = /* @__PURE__ */ deferredLegacySource('xlsx', () => createLegacyXls());
