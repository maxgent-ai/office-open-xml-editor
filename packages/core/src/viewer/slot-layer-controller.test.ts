import { expect, it } from 'vitest';
import { SlotLayerController, type LayerSlot } from './slot-layer-controller';

it('positions a left review rail and previews its marker at the committed scale', () => {
  const marker = { dataset: { ooxmlCommentMarker: '' }, style: {} as CSSStyleDeclaration };
  const markerLayer = { children: [marker], style: {} as CSSStyleDeclaration } as unknown as HTMLDivElement;
  const slot = {
    wrapper: { style: {} as CSSStyleDeclaration },
    canvas: { style: {} as CSSStyleDeclaration },
    textLayer: null,
    highlightLayer: {} as HTMLDivElement,
    elementLayer: null,
    commentMargin: { style: {} as CSSStyleDeclaration },
    commentDecorationLayer: { style: {} as CSSStyleDeclaration },
    renderedScale: 1,
  } as LayerSlot;
  const layers = new SlotLayerController({
    markerLayer: () => markerLayer,
    syncMargin: () => {},
    marginSide: () => 'left',
    marginExtent: () => 80,
    commentsEnabled: () => true,
    previewMargin: () => {},
    disposeMargin: () => {},
    disposeDecoration: () => {},
    redrawOutline: () => {},
    markerRatio: (ratio) => ratio,
    resetMarkerTransform: () => true,
    resetDecorationVisibility: () => false,
  });
  layers.position(0, slot, 24, 300, 200, 500, 12);
  expect(slot.wrapper.style.left).toBe('calc(100px + var(--ooxml-review-origin-x, 0px))');
  expect(slot.commentDecorationLayer!.style.left).toBe('-80px');
  layers.preview(slot, 600, 400, 2);
  expect(marker.style.transform).toBe('translate(-50%,-50%) scale(2)');
});
