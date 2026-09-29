import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { basename, dirname, resolve } from 'node:path';
import { existsSync } from 'node:fs';

const initializedWasm = new WeakMap<object, unknown>();

/** Load and synchronously initialize a wasm-pack `--target web` artifact in
 *  Node. The `--target web` build expects to be `fetch`ed from a URL; in Node
 *  we sidestep that path by reading the .wasm bytes off disk and feeding them
 *  into the generated `initSync` helper. */
export function loadWasmModule<T>(jsModule: T & { initSync: (init: { module: WebAssembly.Module }) => unknown }, wasmPath: string): T {
  const wasmModule = compileWasmModule(wasmPath);
  initializedWasm.set(jsModule as object, jsModule.initSync({ module: wasmModule }));
  return jsModule;
}

/** Compile a shipped parser module once so a realm runtime can use the same
 * immutable module for its initial synchronous instantiation and every forced
 * fresh re-instantiation after a trap. */
export function compileWasmModule(wasmPath: string): WebAssembly.Module {
  return new WebAssembly.Module(readFileSync(wasmPath));
}

/** Defer synchronous filesystem access and compilation until a format is first
 * opened. The Node root entry re-exports all format APIs, so eager module-level
 * compilation would otherwise charge every consumer for all parser binaries. */
export function createLazyWasmModule(
  resolvePath: () => string,
  compile: (wasmPath: string) => WebAssembly.Module = compileWasmModule,
): () => WebAssembly.Module {
  const cache: { value?: WebAssembly.Module } = {};
  return () => cache.value ??= compile(resolvePath());
}

/** Diagnostic linear-memory size for fresh-process benchmarks. */
export function wasmMemoryPages(jsModule: object): number | undefined {
  const initialized = initializedWasm.get(jsModule) as { memory?: WebAssembly.Memory } | undefined;
  return initialized?.memory ? initialized.memory.buffer.byteLength / (64 * 1024) : undefined;
}

/** Resolve a path relative to a workspace-package source file. Used by the
 *  per-format entry points to locate the `.wasm` artifact emitted by
 *  `wasm-pack build --out-dir ../src/wasm`. */
export function resolveWasm(
  metaUrl: string,
  relPath: string,
  workspaceSpecifier?: string,
): string {
  const here = dirname(fileURLToPath(metaUrl));
  const workspacePath = resolve(here, relPath);
  // Source/workspace execution keeps WASM under each format package. The
  // published `@silurus/ooxml/node` bundle places the same emitted assets next
  // to node.mjs, so fall back to that sibling without embedding machine paths.
  if (existsSync(workspacePath)) return workspacePath;
  const shippedPath = resolve(here, basename(relPath));
  if (existsSync(shippedPath) || !workspaceSpecifier) return shippedPath;
  // Node's ESM resolver handles the workspace package export directly; no
  // CommonJS loader capability is needed in a checked OOXML package.
  return fileURLToPath(import.meta.resolve(workspaceSpecifier));
}
