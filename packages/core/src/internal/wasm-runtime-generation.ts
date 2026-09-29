import {
  detachWasmBindgenResource,
  isWasmTrap,
  WasmTrapError,
} from '../worker/wasm-guard.js';
import { RuntimeGeneration } from './runtime-generation.js';

export interface WasmModuleRuntime {
  initSync(input: { module: WebAssembly.Module }): unknown;
  reinit(input: { module_or_path: WebAssembly.Module }): Promise<unknown>;
}

export interface WasmArchiveOpenOptions<TArchive extends object> {
  readonly signal?: AbortSignal;
  readonly abortError?: () => unknown;
  readonly disposeOnAbort: (archive: TArchive) => void;
}

function normalizeTrap(error: unknown): WasmTrapError | null {
  if (!isWasmTrap(error)) return null;
  const detail = error instanceof Error ? error.message : String(error);
  return new WasmTrapError(`WASM parser trapped and was recycled: ${detail}`);
}

/** Realm-local multi-archive owner used by format-owned in-process sessions. */
export class WasmRuntimeGenerationHost<TArchive extends object> {
  private readonly realm: RuntimeGeneration<WasmTrapError>;
  private readonly live = new Set<WasmArchiveHandle<TArchive>>();

  constructor(runtime: WasmModuleRuntime, wasmModule: WebAssembly.Module) {
    this.realm = new RuntimeGeneration(
      () => runtime.initSync({ module: wasmModule }),
      () => runtime.reinit({ module_or_path: wasmModule }),
      normalizeTrap,
    );
    this.realm.onPoison((error) => {
      for (const handle of this.live) handle.poison(error);
      this.live.clear();
    });
  }

  async open(
    create: () => TArchive,
    options?: WasmArchiveOpenOptions<TArchive>,
  ): Promise<WasmArchiveHandle<TArchive>> {
    for (;;) {
      this.throwIfOpenAborted(options);
      await this.awaitReadyOrAbort(options);
      this.throwIfOpenAborted(options);
      const opened = this.realm.tryRunReady(create);
      if (!opened.current) continue;
      const handle = new WasmArchiveHandle(this, opened.value, opened.generation);
      this.live.add(handle);
      if (options?.signal?.aborted) {
        handle.close(options.disposeOnAbort);
        throw options.abortError?.() ?? new DOMException('The operation was aborted', 'AbortError');
      }
      return handle;
    }
  }

  private async awaitReadyOrAbort(
    options: WasmArchiveOpenOptions<TArchive> | undefined,
  ): Promise<void> {
    const signal = options?.signal;
    if (!signal) {
      await this.realm.ensureReady();
      return;
    }
    this.throwIfOpenAborted(options);
    await new Promise<void>((resolve, reject) => {
      const abort = () => reject(
        options?.abortError?.() ?? new DOMException('The operation was aborted', 'AbortError'),
      );
      signal.addEventListener('abort', abort, { once: true });
      this.realm.ensureReady().then(resolve, reject).finally(() => {
        signal.removeEventListener('abort', abort);
      });
    });
  }

  private throwIfOpenAborted(options: WasmArchiveOpenOptions<TArchive> | undefined): void {
    if (!options?.signal?.aborted) return;
    throw options.abortError?.() ?? new DOMException('The operation was aborted', 'AbortError');
  }

  async ensureReady(): Promise<void> {
    await this.realm.ensureReady();
  }

  run<TResult>(handle: WasmArchiveHandle<TArchive>, operation: (archive: TArchive) => TResult): TResult {
    const archive = handle.requireLive(this.realm);
    return this.realm.run(() => operation(archive));
  }

  close(handle: WasmArchiveHandle<TArchive>, free: (archive: TArchive) => void): void {
    const archive = handle.detachForClose();
    this.live.delete(handle);
    if (archive) this.realm.run(() => free(archive));
  }
}

export class WasmArchiveHandle<TArchive extends object> {
  private archive: TArchive | undefined;
  private failure: WasmTrapError | undefined;
  readonly proxy: TArchive;

  constructor(
    private readonly host: WasmRuntimeGenerationHost<TArchive>,
    archive: TArchive,
    private readonly generation: number,
  ) {
    this.archive = archive;
    this.proxy = new Proxy(archive, {
      get: (_target, property) => {
        const value = this.host.run(this, (current) => Reflect.get(current, property, current));
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => this.host.run(this, (current) => {
          const method = Reflect.get(current, property, current) as (...values: unknown[]) => unknown;
          return Reflect.apply(method, current, args);
        });
      },
    }) as TArchive;
  }

  run<TResult>(operation: (archive: TArchive) => TResult): TResult {
    return this.host.run(this, operation);
  }

  close(free: (archive: TArchive) => void): void {
    this.host.close(this, free);
  }

  requireLive(realm: RuntimeGeneration<WasmTrapError>): TArchive {
    if (this.failure) throw this.failure;
    realm.assertCurrent(this.generation);
    if (!this.archive) {
      throw new Error('WASM archive session belongs to a discarded runtime generation');
    }
    return this.archive;
  }

  detachForClose(): TArchive | undefined {
    const archive = this.archive;
    this.archive = undefined;
    return this.failure ? undefined : archive;
  }

  poison(error: WasmTrapError): void {
    const archive = this.archive;
    this.failure = error;
    this.archive = undefined;
    detachWasmBindgenResource(archive);
  }
}
