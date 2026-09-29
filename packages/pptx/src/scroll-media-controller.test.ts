import { expect, it } from 'vitest';
import { PptxScrollMediaController, type MediaSlot } from './scroll-media-controller';

it('retires an interactive handle when its slide leaves the media range', () => {
  let destroyed = 0;
  const slot = {
    mediaInteractive: true,
    presentationGeneration: 4,
    presentationHandle: { destroy: () => { destroyed++; } },
  } as unknown as MediaSlot;
  const slots = new Map([[2, slot]]);
  const controller = new PptxScrollMediaController({
    presentation: () => null, slots: () => slots,
    epoch: () => 0, scale: () => 1, dpr: () => 1,
    width: () => 100, height: () => 100,
    imageResources: () => undefined,
    textSelection: () => false, findActive: () => false,
    mediaEnabled: () => true,
    mediaRange: () => ({ start: 0, end: 0, topIndex: 0, totalHeight: 100 }),
    rangeContains: (range, slide) => slide >= range.start && slide <= range.end,
    shadow: () => false, hyperlinkHandler: () => undefined,
    refreshFindRuns: () => {}, redrawHighlights: () => {},
    commitComments: () => {}, clearTextPreview: () => {}, reportError: () => {},
  });
  controller.sync();
  expect(slot.mediaInteractive).toBe(false);
  expect(slot.presentationHandle).toBeNull();
  expect(slot.presentationGeneration).toBe(5);
  expect(destroyed).toBe(1);
});
