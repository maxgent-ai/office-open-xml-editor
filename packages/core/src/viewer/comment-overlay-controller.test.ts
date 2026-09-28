import { expect, it } from 'vitest';
import { CommentOverlayController, type CommentOverlaySlot } from './comment-overlay-controller';

it('coalesces margin work and ignores a recycled slot before the frame runs', () => {
  const frames: FrameRequestCallback[] = [];
  const drawn: number[] = [];
  const first = { renderedScale: 1 } as CommentOverlaySlot;
  const replacement = { renderedScale: 1 } as CommentOverlaySlot;
  const slots = new Map([[0, first]]);
  const controller = new CommentOverlayController({
    slots: () => slots,
    scale: () => 1,
    destroyed: () => false,
    ownerWindow: () => ({ requestAnimationFrame: (cb: FrameRequestCallback) => {
      frames.push(cb);
      return frames.length;
    } }) as unknown as Window,
    width: () => 100,
    height: () => 100,
    side: () => 'right',
    marginExtent: () => 30,
    connectorOptions: () => undefined,
    runtime: () => null,
    redrawComments: (unit) => drawn.push(unit),
  });

  controller.schedule(0, first, true);
  controller.schedule(0, first, false);
  expect(frames).toHaveLength(1);
  frames.shift()!(0);
  expect(drawn).toEqual([0]);

  controller.schedule(0, first);
  slots.set(0, replacement);
  frames.shift()!(0);
  expect(drawn).toEqual([0]);
  controller.destroy();
});
