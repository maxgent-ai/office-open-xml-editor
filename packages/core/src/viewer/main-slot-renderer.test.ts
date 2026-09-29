import { expect, it } from 'vitest';
import { MainSlotRenderer, type MainSlot } from './main-slot-renderer';

it('does not commit text overlays from a superseded main-thread render', async () => {
  let epoch = 0;
  let resolvePaint!: () => void;
  let committed = 0;
  const slot = {
    canvas: {} as HTMLCanvasElement,
    wrapper: {} as HTMLDivElement,
    dispatcher: { isCurrent: () => true } as unknown as MainSlot['dispatcher'],
    renderedScale: -1,
    index: 0,
  };
  const slots = new Map([[0, slot]]);
  const renderer = new MainSlotRenderer({
    slots: () => slots,
    epoch: () => epoch,
    scale: () => epoch + 1,
    token: () => 0,
    nextToken: () => 0,
    slotIndex: (live) => live.index,
    wantRuns: () => true,
    render: async () => await new Promise<void>((resolve) => { resolvePaint = resolve; }),
    commitRuns: () => { committed++; },
    shadow: () => false,
    reportError: () => {},
  });
  const pending = renderer.render(0, slot, 100, 1, 0, slot.dispatcher, 1, true);
  epoch = 1;
  resolvePaint();
  await pending;
  expect(committed).toBe(0);
  expect(slot.renderedScale).toBe(-1);
});
