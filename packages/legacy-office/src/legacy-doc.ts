import { resolveBundledAssetUrl } from '@silurus/ooxml-core/internal/bundled-asset-url';
import type { ModelSource } from '@silurus/ooxml-core';
import { createLegacySource, deferredLegacySource, type LegacySourceOptions } from './legacy-source.js';
import defaultWasmUrl from './wasm-direct-doc/legacy_doc_direct_bg.wasm?url';
import defaultModuleUrl from './legacy-doc-source-module.ts?url';

export type LegacyDocOptions = LegacySourceOptions;

/**
 * Create a model source for legacy Word 97-2003 (.doc) input with custom
 * options, for example `wasmUrl` / `moduleUrl` when an asset pipeline
 * relocates the bundled files. Pass it as the `legacyDoc` option of the DOCX
 * loaders and viewers. Creating it fetches nothing; the source module and its
 * WASM load only when a .doc input is claimed.
 */
export function createLegacyDoc(options: LegacyDocOptions = {}): ModelSource<'docx'> {
  if (typeof options !== 'object' || options === null || Array.isArray(options))
    throw new TypeError('legacy DOC source options must be an object');
  return createLegacySource<'docx'>('doc', options, {
    wasmUrl: resolveBundledAssetUrl(defaultWasmUrl, import.meta.url),
    moduleUrl: resolveBundledAssetUrl(defaultModuleUrl, import.meta.url),
  });
}

/**
 * Legacy Word 97-2003 (.doc) reader with the bundled asset URLs:
 * `import { legacyDoc } from '@silurus/ooxml-legacy-office/doc'`, then pass
 * `{ legacyDoc }` to the DOCX loaders and viewers. Importing it fetches and
 * instantiates nothing; the bundled URLs resolve on first use, and the source
 * module and its WASM load only when a .doc input is claimed. Use
 * {@link createLegacyDoc} for custom asset URLs or an input size limit.
 */
export const legacyDoc: ModelSource<'docx'> = /* @__PURE__ */ deferredLegacySource('docx', () => createLegacyDoc());
