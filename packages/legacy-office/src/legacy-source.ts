/**
 * Shared calling-realm half of the legacy Office model sources. The legacy
 * readers plug into the DOCX / XLSX / PPTX renderers only through the generic
 * `ModelSource` contract: each factory claims its own CFB family from the raw
 * bytes and names a self-contained source module plus its WASM URL.
 */
import {
  type ModelSource,
  type ModelSourceModuleDescriptor,
  type ModelSourceTarget,
} from '@silurus/ooxml-core';
import { cfbDirectoryNames } from '@silurus/ooxml-core/internal/cfb-directory-names';
import { assertCompanionRuntime, MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml/internal/companion';
import { COMPANION_VERSION } from '../../../src/companion-version.js';

import { readLegacyMaxInputBytes } from './source-module-config.js';

export type LegacyFamily = 'doc' | 'xls' | 'ppt';

/** Options accepted by every legacy source factory. */
export interface LegacySourceOptions {
  /** Absolute URL of the reader's WASM binary. Defaults to the emitted asset. */
  readonly wasmUrl?: string;
  /**
   * Absolute URL of the reader's self-contained ES source module. Defaults to
   * the emitted asset; override it when an asset pipeline relocates files.
   */
  readonly moduleUrl?: string;
  /** Optional caller limit on input bytes; omitted means no whole-file gate. */
  readonly maxInputBytes?: number;
}

const URL_PROTOCOLS: Readonly<Record<'wasmUrl' | 'moduleUrl', ReadonlySet<string>>> = {
  wasmUrl: new Set(['http:', 'https:', 'file:', 'blob:', 'data:']),
  moduleUrl: new Set(['http:', 'https:', 'file:', 'blob:']),
};

const TARGETS: Readonly<Record<LegacyFamily, ModelSourceTarget>> = {
  doc: 'docx',
  xls: 'xlsx',
  ppt: 'pptx',
};

/**
 * [MS-CFB] stream names that identify each binary family: WordDocument
 * ([MS-DOC] 2.1.1), Workbook/Book ([MS-XLS] 2.1.2) and PowerPoint Document
 * ([MS-PPT] 2.1.1). A container carrying more than one family (for example
 * through embedded objects) is ambiguous and is not claimed.
 */
function cfbFamily(names: ReadonlySet<string>): LegacyFamily | null {
  const families: LegacyFamily[] = [];
  if (names.has('WordDocument')) families.push('doc');
  if (names.has('Workbook') || names.has('Book')) families.push('xls');
  if (names.has('PowerPoint Document')) families.push('ppt');
  return families.length === 1 ? families[0] as LegacyFamily : null;
}

export function createLegacySource<T extends ModelSourceTarget>(
  family: LegacyFamily,
  options: LegacySourceOptions,
  defaults: Readonly<{ wasmUrl: string; moduleUrl: string }>,
): ModelSource<T> {
  // The installed viewer owns protocol and transport. Pin our compiled
  // expectation too: copying its current protocol would hide a stale add-on.
  assertCompanionRuntime(COMPANION_VERSION, 'ooxml-model-source-module/v2');
  const label = `legacy ${family.toUpperCase()} source`;
  if (typeof options !== 'object' || options === null) {
    throw new TypeError(`${label} options must be an object`);
  }
  const wasmUrl = absoluteUrl(options.wasmUrl ?? defaults.wasmUrl, 'wasmUrl', label);
  const moduleUrl = absoluteUrl(options.moduleUrl ?? defaults.moduleUrl, 'moduleUrl', label);
  const maxInputBytes = readLegacyMaxInputBytes(options.maxInputBytes, label);
  const target = TARGETS[family] as T;
  const module: ModelSourceModuleDescriptor = Object.freeze({
    protocol: MODEL_SOURCE_MODULE_PROTOCOL,
    target,
    moduleUrl,
    config: Object.freeze({ wasmUrl, ...(maxInputBytes === undefined ? {} : { maxInputBytes }) }),
  });
  return Object.freeze({
    target,
    claim(bytes: Uint8Array): boolean {
      const names = cfbDirectoryNames(bytes);
      // Encrypted packages stay on the OOXML path, which reports them.
      if (!names || names.has('EncryptionInfo') || cfbFamily(names) !== family) return false;
      if (maxInputBytes !== undefined && bytes.byteLength > maxInputBytes) {
        throw new RangeError(`${label} input exceeds the configured size limit`);
      }
      return true;
    },
    beginLoad() {
      // The calling realm retains no transfer or native handle. The opened
      // archive belongs to the source module and is closed by its owner;
      // release is therefore an idempotent no-op even on an open failure.
      return { module, release() {} };
    },
  });
}

/**
 * A source with the bundled defaults that can be exported as a constant. The
 * real source (companion check and asset URL resolution) is created on the
 * first `claim()` or `beginLoad()`, never at import time, so importing the
 * entry fetches, instantiates and resolves nothing.
 */
export function deferredLegacySource<T extends ModelSourceTarget>(
  target: T,
  create: () => ModelSource<T>,
): ModelSource<T> {
  let source: ModelSource<T> | undefined;
  const resolve = (): ModelSource<T> => (source ??= create());
  return Object.freeze({
    target,
    claim: (bytes: Uint8Array): boolean => resolve().claim(bytes),
    beginLoad: () => resolve().beginLoad(),
  });
}

function absoluteUrl(value: unknown, field: 'wasmUrl' | 'moduleUrl', label: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.trim() !== value) {
    throw new TypeError(`${label} ${field} must be a nonempty absolute URL`);
  }
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new TypeError(`${label} ${field} must be a nonempty absolute URL`);
  }
  if (!URL_PROTOCOLS[field].has(url.protocol)) {
    throw new TypeError(`${label} ${field} protocol is unsupported`);
  }
  return value;
}
