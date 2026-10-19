import { describe, expect, it, vi } from 'vitest';
import {
  type ModelSource,
  type ModelSourceTarget,
} from '@silurus/ooxml-core';
import {
  MODEL_SOURCE_MODULE_PROTOCOL,
  validateModelSourceModuleDescriptor,
} from '@silurus/ooxml-core/internal/model-source';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';

// Creating, claiming and describing a source must never load native code:
// the glue and source modules may be imported only by the archive realm.
const loaded = vi.hoisted(() => [] as string[]);
vi.mock('./wasm-direct-doc/legacy_doc_direct.js', () => { loaded.push('doc glue'); return {}; });
vi.mock('./wasm-direct-xls/legacy_xls_direct.js', () => { loaded.push('xls glue'); return {}; });
vi.mock('./wasm-direct-ppt/legacy_ppt_direct.js', () => { loaded.push('ppt glue'); return {}; });
vi.mock('./legacy-doc-source-module.ts', () => { loaded.push('doc module'); return {}; });
vi.mock('./legacy-xls-source-module.ts', () => { loaded.push('xls module'); return {}; });
vi.mock('./legacy-ppt-source-module.ts', () => { loaded.push('ppt module'); return {}; });

import { createLegacyDoc, legacyDoc } from './legacy-doc.js';
import { createLegacyPpt, legacyPpt } from './legacy-ppt.js';
import { createLegacySource, deferredLegacySource, type LegacyFamily } from './legacy-source.js';
import { createLegacyXls, legacyXls } from './legacy-xls.js';
import { buildStoredZip } from './zip-fixture.js';

const cfb = (...streams: string[]) => new Uint8Array(buildCfbFixture(['Root Entry', ...streams]));
const WASM = 'https://cdn.example.test/legacy.wasm';
const MODULE = 'https://cdn.example.test/legacy-source-module.js';

const factories = [
  { family: 'doc', target: 'docx', create: createLegacyDoc, constant: legacyDoc, streams: [['WordDocument']] },
  { family: 'xls', target: 'xlsx', create: createLegacyXls, constant: legacyXls, streams: [['Workbook'], ['Book']] },
  { family: 'ppt', target: 'pptx', create: createLegacyPpt, constant: legacyPpt, streams: [['PowerPoint Document']] },
] as const satisfies ReadonlyArray<{
  family: LegacyFamily;
  target: ModelSourceTarget;
  create: (options?: { wasmUrl?: string; moduleUrl?: string; maxInputBytes?: number }) => ModelSource;
  constant: ModelSource;
  streams: ReadonlyArray<readonly string[]>;
}>;

describe.each(factories)('legacy $family default source constant', ({ target, create, constant, streams }) => {
  it('is a frozen source for its own target that describes the bundled defaults', () => {
    expect(constant.target).toBe(target);
    expect(Object.isFrozen(constant)).toBe(true);
    expect(constant.beginLoad().module).toEqual(create().beginLoad().module);
    expect(constant.claim(cfb(...streams[0]))).toBe(true);
    expect(constant.claim(new Uint8Array([1, 2, 3]))).toBe(false);
  });
});

describe('deferredLegacySource', () => {
  it('creates the real source once, on first use rather than at definition', () => {
    const real = { target: 'docx', claim: vi.fn(() => true), beginLoad: vi.fn() } as unknown as ModelSource<'docx'>;
    const create = vi.fn(() => real);
    const deferred = deferredLegacySource('docx', create);
    expect(create).not.toHaveBeenCalled();
    expect(deferred.target).toBe('docx');
    expect(deferred.claim(new Uint8Array())).toBe(true);
    deferred.beginLoad();
    expect(create).toHaveBeenCalledOnce();
    expect(real.beginLoad).toHaveBeenCalledOnce();
  });

  it('surfaces a creation failure at first use and retries on the next', () => {
    const failure = new Error('stale companion');
    const create = vi.fn(() => { throw failure; });
    const deferred = deferredLegacySource('xlsx', create);
    expect(() => deferred.claim(new Uint8Array())).toThrow(failure);
    expect(() => deferred.beginLoad()).toThrow(failure);
    expect(create).toHaveBeenCalledTimes(2);
  });
});

describe.each(factories)('legacy $family source factory', ({ family, target, create, streams }) => {
  it('describes its own target with frozen, admissible default asset URLs', () => {
    const source = create();
    expect(source.target).toBe(target);
    expect(Object.isFrozen(source)).toBe(true);
    const load = source.beginLoad();
    const { module } = load;
    expect(Object.isFrozen(module)).toBe(true);
    expect(Object.isFrozen(module.config)).toBe(true);
    expect(module).toEqual({
      protocol: MODEL_SOURCE_MODULE_PROTOCOL,
      target,
      moduleUrl: expect.any(String),
      config: { wasmUrl: expect.any(String) },
    });
    expect(new URL(module.moduleUrl).pathname).toContain(`legacy-${family}-source-module`);
    expect(new URL(String(module.config.wasmUrl)).pathname)
      .toContain(`/wasm-direct-${family}/legacy_${family}_direct_bg.wasm`);
    // The core loader admits the descriptor for this target and no other.
    expect(validateModelSourceModuleDescriptor(module, target)).toEqual(module);
    expect(load.transfer).toBeUndefined();
    expect(() => load.release()).not.toThrow();
    expect(() => load.release()).not.toThrow();
  });

  it('passes explicit URLs and budget through verbatim', () => {
    const { module } = create({ wasmUrl: WASM, moduleUrl: MODULE, maxInputBytes: 4096 }).beginLoad();
    expect(module).toEqual({
      protocol: MODEL_SOURCE_MODULE_PROTOCOL,
      target,
      moduleUrl: MODULE,
      config: { wasmUrl: WASM, maxInputBytes: 4096 },
    });
  });

  it('claims only an unencrypted CFB whose single binary family is its own', () => {
    const source = create();
    for (const own of streams) expect(source.claim(cfb(...own))).toBe(true);
    const others = factories.filter((other) => other.family !== family);
    for (const other of others) {
      for (const foreign of other.streams) expect(source.claim(cfb(...foreign))).toBe(false);
      // Embedded objects can make the family ambiguous; do not guess.
      expect(source.claim(cfb(...streams[0], ...other.streams[0]))).toBe(false);
    }
    // Encrypted packages stay on the OOXML path, which reports them.
    expect(source.claim(cfb('EncryptionInfo', 'EncryptedPackage', ...streams[0]))).toBe(false);
    expect(source.claim(cfb())).toBe(false);
    expect(source.claim(buildStoredZip({ '[Content_Types].xml': '<Types/>' }))).toBe(false);
    expect(source.claim(new Uint8Array([1, 2, 3]))).toBe(false);
  });

  it('rejects oversize claimed input and ignores oversize unclaimed input', () => {
    const own = cfb(...streams[0]);
    const source = create({ maxInputBytes: own.byteLength - 1 });
    expect(() => source.claim(own)).toThrow(RangeError);
    expect(create({ maxInputBytes: own.byteLength }).claim(own)).toBe(true);
    const foreign = factories.find((other) => other.family !== family)!.streams[0];
    expect(source.claim(cfb(...foreign))).toBe(false);
    expect(source.claim(buildStoredZip({ big: new Uint8Array(own.byteLength) }))).toBe(false);
  });
});

describe('createLegacySource option validation', () => {
  const defaults = { wasmUrl: WASM, moduleUrl: MODULE };
  const make = (options: unknown) => createLegacySource('doc', options as never, defaults);

  it('fails closed on non-object options', () => {
    for (const options of [null, 'https://example.test/doc.wasm', 1]) {
      expect(() => make(options)).toThrow(TypeError);
    }
  });

  it('admits absolute URLs of the supported protocols; data: only for WASM', () => {
    for (const wasmUrl of [
      'http://example.test/doc.wasm',
      'file:///opt/assets/doc.wasm',
      'blob:https://example.test/1234',
      'data:application/wasm;base64,AGFzbQ==',
    ]) {
      expect(make({ wasmUrl }).beginLoad().module.config.wasmUrl).toBe(wasmUrl);
    }
    for (const moduleUrl of ['http://example.test/m.js', 'file:///opt/m.js', 'blob:https://example.test/1234']) {
      expect(make({ moduleUrl }).beginLoad().module.moduleUrl).toBe(moduleUrl);
    }
    expect(() => make({ moduleUrl: 'data:text/javascript,export%20default%201' })).toThrow(/protocol/);
  });

  it.each(['wasmUrl', 'moduleUrl'] as const)('rejects a malformed %s', (field) => {
    for (const value of ['', './relative', ' https://example.test/x', 'https://example.test/x\n', 'javascript:alert(1)', 42]) {
      expect(() => make({ [field]: value })).toThrow(TypeError);
    }
  });

  it('accepts an optional positive safe integer without a library size ceiling', () => {
    for (const maxInputBytes of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Number.MAX_SAFE_INTEGER + 1]) {
      expect(() => make({ maxInputBytes })).toThrow(RangeError);
    }
    expect(make({}).beginLoad().module.config).not.toHaveProperty('maxInputBytes');
    expect(make({ maxInputBytes: 1 }).beginLoad().module.config.maxInputBytes).toBe(1);
    expect(make({ maxInputBytes: Number.MAX_SAFE_INTEGER }).beginLoad().module.config.maxInputBytes)
      .toBe(Number.MAX_SAFE_INTEGER);
  });

  it('never loaded a glue or source module while sources were created and used', () => {
    expect(loaded).toEqual([]);
  });
});
