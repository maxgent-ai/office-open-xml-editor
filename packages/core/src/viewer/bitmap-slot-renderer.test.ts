import { expect, it } from 'vitest';
import { BitmapSlotRenderer } from './bitmap-slot-renderer';
import type { BitmapSlot } from './bitmap-slot-renderer';

it('closes a stale bitmap and follows the current epoch without retrying a plain failure', async () => {
  let epoch = 0;
  let dispatches = 0;
  let resolveFirst!: (bitmap: ImageBitmap) => void;
  let committed = 0;
  let closed = 0;
  const dispatcher = {
    begin: () => 1,
    isCurrent: () => true,
  };
  const slot = {
    canvas: {} as HTMLCanvasElement,
    dispatcher: dispatcher as unknown as BitmapSlot['dispatcher'],
    renderedScale: -1, renderedUnit: 0,
  };
  const slots = new Map([[0, slot]]);
  const inFlight = new Set<number>();
  const renderer = new BitmapSlotRenderer({
    slots: () => slots, inFlight: () => inFlight,
    epoch: () => epoch, destroyed: () => false,
    scale: () => epoch + 1, width: () => 100, dpr: () => 1,
    slotIndex: (live) => live.renderedUnit,
    token: () => 0, nextToken: () => 0, canRetry: () => true,
    wantRuns: () => false,
    render: async () => {
      dispatches++;
      if (dispatches === 1) return await new Promise<ImageBitmap>((resolve) => { resolveFirst = resolve; });
      return { close: () => { closed++; } } as ImageBitmap;
    },
    commitBitmap: () => { committed++; return true; },
    commitRuns: () => {}, reportError: () => {},
  });
  const result = renderer.render(0, slot, 100, 1, 1);
  epoch = 1;
  resolveFirst({ close: () => { closed++; } } as ImageBitmap);
  await result;
  expect([dispatches, committed, closed, slot.renderedScale]).toEqual([2, 1, 1, 2]);

  const failed = new BitmapSlotRenderer({
    slots: () => slots, inFlight: () => inFlight,
    epoch: () => epoch, destroyed: () => false,
    scale: () => 2, width: () => 100, dpr: () => 1,
    slotIndex: (live) => live.renderedUnit,
    token: () => 0, nextToken: () => 0, canRetry: () => true,
    wantRuns: () => false,
    render: async () => { dispatches++; throw new Error('render failed'); },
    commitBitmap: () => true, commitRuns: () => {}, reportError: () => {},
  });
  await failed.render(0, slot, 100, 1, 2);
  expect(dispatches).toBe(3);
});
