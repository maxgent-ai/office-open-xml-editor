import { renderCanvasElementOutline } from '../internal/canvas-viewer-mechanics';
import { previewSlotHost, resetSlotHost, type SlotHost } from './slot-scroller';

export interface LayerSlot extends SlotHost {
  elementLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentDecorationLayer: HTMLDivElement | null;
  renderedScale: number;
}

export interface LayerHooks<Slot extends LayerSlot> {
  markerLayer(slot: Slot): HTMLDivElement | null;
  syncMargin(margin: HTMLDivElement | null): void;
  marginSide(): 'left' | 'right';
  marginExtent(): number;
  commentsEnabled(): boolean;
  previewMargin(margin: HTMLDivElement, ratio: number): void;
  disposeMargin(margin: HTMLDivElement): void;
  disposeDecoration(layer: HTMLDivElement): void;
  redrawOutline(unit: number, slot: Slot): void;
  markerRatio(ratio: number): number;
  resetMarkerTransform(): boolean;
  resetDecorationVisibility(): boolean;
}

/** Positions and previews the shared canvas, comment, and outline stack. */
export class SlotLayerController<Slot extends LayerSlot> {
  constructor(private readonly hooks: LayerHooks<Slot>) {}

  position(unit: number, slot: Slot, top: number, width: number, height: number,
    viewportWidth: number, padLeft: number): void {
    slot.wrapper.style.top = `${top}px`;
    slot.wrapper.style.width = `${width}px`;
    slot.wrapper.style.height = `${height}px`;
    this.hooks.syncMargin(slot.commentMargin);
    if (slot.commentDecorationLayer) {
      const extent = this.hooks.marginExtent();
      slot.commentDecorationLayer.style.left = this.hooks.marginSide() === 'left'
        ? `${-extent}px` : '0px';
      slot.commentDecorationLayer.style.width = `${width + extent}px`;
      slot.commentDecorationLayer.style.height = `${height}px`;
    }
    this.hooks.redrawOutline(unit, slot);
    const authoredLeft = Math.max(padLeft, (viewportWidth - width) / 2);
    slot.wrapper.style.left = this.hooks.marginSide() === 'left' && this.hooks.commentsEnabled()
      ? `calc(${authoredLeft}px + var(--ooxml-review-origin-x, 0px))`
      : `${authoredLeft}px`;
  }

  preview(slot: Slot, width: number, height: number, scale: number): void {
    const ratio = previewSlotHost(slot, width, height, scale);
    const markerLayer = this.hooks.markerLayer(slot);
    if (ratio !== null) {
      if (slot.commentMargin) this.hooks.previewMargin(slot.commentMargin, ratio);
      for (const marker of markerLayer?.children ?? []) {
        if ((marker as HTMLElement).dataset.ooxmlCommentMarker === undefined) continue;
        (marker as HTMLElement).style.transform =
          `translate(-50%,-50%) scale(${this.hooks.markerRatio(ratio)})`;
      }
    }
    const visibility = ratio === null ? 'hidden' : '';
    if (markerLayer) markerLayer.style.visibility = visibility;
    if (slot.commentMargin) slot.commentMargin.style.visibility = visibility;
    if (slot.commentDecorationLayer) slot.commentDecorationLayer.style.visibility = visibility;
  }

  reset(slot: Slot): void {
    resetSlotHost(slot);
    const markerLayer = this.hooks.markerLayer(slot);
    if (markerLayer) {
      markerLayer.replaceChildren();
      if (this.hooks.resetMarkerTransform()) {
        markerLayer.style.transform = '';
        markerLayer.style.transformOrigin = '';
      }
      markerLayer.style.visibility = '';
    }
    if (slot.commentMargin) {
      this.hooks.disposeMargin(slot.commentMargin);
      slot.commentMargin.style.visibility = '';
    }
    if (slot.commentDecorationLayer) {
      this.hooks.disposeDecoration(slot.commentDecorationLayer);
      if (this.hooks.resetDecorationVisibility()) slot.commentDecorationLayer.style.visibility = '';
    }
    renderCanvasElementOutline(slot.elementLayer, null);
  }
}
