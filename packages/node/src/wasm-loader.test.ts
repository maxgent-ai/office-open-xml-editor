import { describe, expect, it, vi } from 'vitest';
import { createLazyWasmModule, resolveWasm } from './wasm-loader.ts';

describe('createLazyWasmModule', () => {
  it('resolves and compiles only on first use, then reuses the immutable module', () => {
    const wasmModule = {} as WebAssembly.Module;
    const resolvePath = vi.fn(() => '/parser.wasm');
    const compile = vi.fn(() => wasmModule);
    const getModule = createLazyWasmModule(resolvePath, compile);

    expect(resolvePath).not.toHaveBeenCalled();
    expect(compile).not.toHaveBeenCalled();

    expect(getModule()).toBe(wasmModule);
    expect(getModule()).toBe(wasmModule);
    expect(resolvePath).toHaveBeenCalledOnce();
    expect(compile).toHaveBeenCalledOnce();
    expect(compile).toHaveBeenCalledWith('/parser.wasm');
  });
});

it('resolves the workspace WASM export through ESM when the local sibling is absent', () => {
  expect(resolveWasm(import.meta.url, 'missing-parser.wasm', '@silurus/ooxml-docx/wasm-binary'))
    .toMatch(/docx_parser_bg\.wasm$/);
});
