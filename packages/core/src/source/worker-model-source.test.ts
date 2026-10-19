import { describe, expect, it, vi } from 'vitest';
import { copyingCursorArchive, runModelSourceOperation } from './worker-model-source';

describe('copyingCursorArchive', () => {
  it('copies the cursor bytes and keeps the archive as `this`', () => {
    const shared = new Uint8Array([9, 1, 2, 3, 9]);
    const archive = {
      offset: 1,
      pull(length: number) { return shared.subarray(this.offset, this.offset + length); },
      other() { return this.offset; },
    };
    const wrapped = copyingCursorArchive(archive, 'pull');
    const bytes = wrapped.pull(3);
    expect([...bytes]).toEqual([1, 2, 3]);
    expect(bytes.buffer).not.toBe(shared.buffer);
    expect(bytes.byteOffset).toBe(0);
    expect(bytes.buffer.byteLength).toBe(3);
    expect(wrapped.other()).toBe(1);
  });
});

describe('runModelSourceOperation', () => {
  it('closes the source only on a WASM trap, then rethrows', () => {
    const onTrap = vi.fn(() => { throw new Error('close failed'); });
    const trap = new WebAssembly.RuntimeError('unreachable');
    expect(() => runModelSourceOperation(() => { throw trap; }, onTrap)).toThrow(trap);
    expect(onTrap).toHaveBeenCalledTimes(1);
    const ordinary = new Error('ordinary');
    expect(() => runModelSourceOperation(() => { throw ordinary; }, onTrap)).toThrow(ordinary);
    expect(onTrap).toHaveBeenCalledTimes(1);
    expect(runModelSourceOperation(() => 7, onTrap)).toBe(7);
  });
});
