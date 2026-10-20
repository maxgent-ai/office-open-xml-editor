import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, deferredLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-ppt/legacy_ppt_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-ppt-source-module.ts?url';

export type LegacyPptOptions = LegacySourceOptions;

/**
 * Create a model source for legacy PowerPoint 97-2003 (.ppt) input with custom
 * options, for example `wasmUrl` / `moduleUrl` when an asset pipeline
 * relocates the bundled files. Pass it as the `legacyPpt` option of the PPTX
 * loaders and viewers. Creating it fetches nothing; the source module and its
 * WASM load only when a .ppt input is claimed.
 */
export function createLegacyPpt(options: LegacyPptOptions = {}): ModelSource<'pptx'> {
  return createLegacySource<'pptx'>('ppt', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * Legacy PowerPoint 97-2003 (.ppt) reader with the bundled asset URLs:
 * `import { legacyPpt } from '@silurus/ooxml-legacy-office/ppt'`, then pass
 * `{ legacyPpt }` to the PPTX loaders and viewers. Importing it fetches and
 * instantiates nothing; the bundled URLs resolve on first use, and the source
 * module and its WASM load only when a .ppt input is claimed. Use
 * {@link createLegacyPpt} for custom asset URLs or an input size limit.
 */
export const legacyPpt: ModelSource<'pptx'> = /* @__PURE__ */ deferredLegacySource('pptx', () => createLegacyPpt());
