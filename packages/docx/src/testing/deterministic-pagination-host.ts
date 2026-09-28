import { vi } from 'vitest';

/** Force the default pagination scheduler to yield after every suspension.
 * Callers restore the clock spy and MessageChannel stub in afterEach.
 */
export function installDeterministicPaginationHost(): () => number {
  let clock = 0;
  let yields = 0;
  vi.spyOn(globalThis.performance, 'now').mockImplementation(() => {
    clock += 17;
    return clock;
  });
  class TestMessageChannel {
    readonly port1 = {
      onmessage: null as (() => void) | null,
      close: () => undefined,
    };
    readonly port2 = {
      close: () => undefined,
      postMessage: () => queueMicrotask(() => {
        yields++;
        this.port1.onmessage?.();
      }),
    };
  }
  vi.stubGlobal('MessageChannel', TestMessageChannel);
  return () => yields;
}
