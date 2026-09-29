import { readFileSync } from 'node:fs';
import { describe, expect, it, vi } from 'vitest';
import { WasmTrapError } from '@silurus/ooxml-core/worker';
import {
  WasmRuntimeGenerationHost,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';

type Archive = { value(): number; free(): void };

function runtimeHost() {
  const initSync = vi.fn();
  const reinit = vi.fn(async () => undefined);
  const host = new WasmRuntimeGenerationHost<Archive>(
    { initSync, reinit },
    new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0])),
  );
  return { host, initSync, reinit };
}

describe('WasmRuntimeGenerationHost', () => {
  it('initializes once and permits concurrent archive handles in one generation', async () => {
    const { host, initSync, reinit } = runtimeHost();
    const first = await host.open(() => ({ value: () => 1, free: vi.fn() }));
    const second = await host.open(() => ({ value: () => 2, free: vi.fn() }));

    expect(first.proxy.value()).toBe(1);
    expect(second.proxy.value()).toBe(2);
    expect(initSync).toHaveBeenCalledOnce();
    expect(reinit).not.toHaveBeenCalled();
  });

  it('poisons every live handle and reinitializes once for concurrent next opens', async () => {
    const { host, reinit } = runtimeHost();
    const trappedDetach = vi.fn(() => 1);
    const trapped = await host.open(() => ({
      value: () => { throw new WebAssembly.RuntimeError('unreachable'); },
      free: vi.fn(),
      __destroy_into_raw: trappedDetach,
    }));
    const siblingFree = vi.fn();
    const siblingDetach = vi.fn(() => 2);
    const sibling = await host.open(() => ({
      value: () => 2,
      free: siblingFree,
      __destroy_into_raw: siblingDetach,
    }));

    expect(() => trapped.proxy.value()).toThrow(WasmTrapError);
    expect(() => sibling.proxy.value()).toThrow(WasmTrapError);
    expect(trappedDetach).toHaveBeenCalledOnce();
    expect(siblingDetach).toHaveBeenCalledOnce();
    sibling.close((archive) => archive.free());
    expect(siblingFree).not.toHaveBeenCalled();

    const [next, concurrent] = await Promise.all([
      host.open(() => ({ value: () => 3, free: vi.fn() })),
      host.open(() => ({ value: () => 4, free: vi.fn() })),
    ]);
    expect(reinit).toHaveBeenCalledOnce();
    expect(next.proxy.value()).toBe(3);
    expect(concurrent.proxy.value()).toBe(4);
  });

  it('uses the generated JS-only detach before reinit so a stale finalizer cannot own the fresh archive', async () => {
    type GeneratedArchive = {
      __destroy_into_raw(): number;
      free(): void;
      parse(): Uint8Array;
    };
    type GeneratedArchiveConstructor = new (
      bytes: Uint8Array,
      maxEntry: bigint | null,
      maxTotal: bigint | null,
      maxEntries: bigint | null,
    ) => GeneratedArchive;

    const wasmModule = new WebAssembly.Module(readFileSync(new URL(
      '../../xlsx/src/wasm/xlsx_parser_bg.wasm',
      import.meta.url,
    )));
    const bytes = readFileSync(new URL('../../xlsx/public/demo/sample-1.xlsx', import.meta.url));
    const Archive = (xlsxWasm as unknown as { XlsxArchive: GeneratedArchiveConstructor }).XlsxArchive;
    const detach = vi.spyOn(Archive.prototype, '__destroy_into_raw');
    const host = new WasmRuntimeGenerationHost<GeneratedArchive>(
      xlsxWasm as unknown as WasmModuleRuntime,
      wasmModule,
    );
    let recovered: import('@silurus/ooxml-core/internal/wasm-runtime-generation')
      .WasmArchiveHandle<GeneratedArchive> | undefined;

    try {
      const stale = await host.open(() => new Archive(bytes, null, null, null));
      expect(() => stale.run(() => { throw new RangeError('synthetic trap'); }))
        .toThrow(WasmTrapError);
      expect(detach).toHaveBeenCalledOnce();

      recovered = await host.open(() => new Archive(bytes, null, null, null));
      expect(recovered.run((archive) => archive.parse()).byteLength).toBeGreaterThan(0);
    } finally {
      recovered?.close((archive) => archive.free());
      detach.mockRestore();
    }
  });

  it('moves a sibling concurrent open to a fresh generation when create traps', async () => {
    const { host, reinit } = runtimeHost();
    const firstCreate = vi.fn((): Archive => {
      throw new WebAssembly.RuntimeError('unreachable');
    });
    const siblingCreate = vi.fn(() => ({ value: () => 2, free: vi.fn() }));

    const [first, sibling] = await Promise.allSettled([
      host.open(firstCreate),
      host.open(siblingCreate),
    ]);

    expect(first.status).toBe('rejected');
    if (first.status === 'rejected') expect(first.reason).toBeInstanceOf(WasmTrapError);
    expect(sibling.status).toBe('fulfilled');
    if (sibling.status === 'fulfilled') expect(sibling.value.proxy.value()).toBe(2);
    expect(firstCreate).toHaveBeenCalledOnce();
    expect(siblingCreate).toHaveBeenCalledOnce();
    expect(reinit).toHaveBeenCalledOnce();
  });

  it('observes aborts while a recycled runtime is initializing before constructing an archive', async () => {
    const reinitGate = deferred<void>();
    const runtime: WasmModuleRuntime = {
      initSync: vi.fn(),
      reinit: vi.fn(() => reinitGate.promise),
    };
    const host = new WasmRuntimeGenerationHost<Archive>(
      runtime,
      new WebAssembly.Module(new Uint8Array([0, 97, 115, 109, 1, 0, 0, 0])),
    );
    const trapped = await host.open(() => ({
      value: () => { throw new WebAssembly.RuntimeError('unreachable'); },
      free: vi.fn(),
    }));
    expect(() => trapped.proxy.value()).toThrow(WasmTrapError);

    const controller = new AbortController();
    const create = vi.fn(() => ({ value: () => 2, free: vi.fn() }));
    const opening = host.open(create, {
      signal: controller.signal,
      abortError: () => Object.assign(new Error('format session was aborted'), {
        name: 'AbortError',
      }),
      disposeOnAbort: (archive) => archive.free(),
    });
    controller.abort();

    await expect(opening).rejects.toMatchObject({ name: 'AbortError' });
    expect(create).not.toHaveBeenCalled();

    // Cancellation is caller-local: the shared realm recovery keeps running
    // and can serve a later caller once it settles.
    reinitGate.resolve();
    const recovered = await host.open(() => ({ value: () => 3, free: vi.fn() }));
    expect(recovered.proxy.value()).toBe(3);
  });

  it('frees a healthy archive exactly once', async () => {
    const { host } = runtimeHost();
    const free = vi.fn();
    const handle = await host.open(() => ({ value: () => 1, free }));

    handle.close((archive) => archive.free());
    handle.close((archive) => archive.free());
    expect(free).toHaveBeenCalledOnce();
    expect(() => handle.proxy.value()).toThrow(/discarded runtime generation/);
  });
});

function deferred<T>(): { promise: Promise<T>; resolve(value?: T): void } {
  let resolvePromise: (value: T | PromiseLike<T>) => void = () => undefined;
  const promise = new Promise<T>((resolve) => { resolvePromise = resolve; });
  return {
    promise,
    resolve: (value?: T) => resolvePromise(value as T),
  };
}
