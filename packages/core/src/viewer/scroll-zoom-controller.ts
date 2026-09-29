import { anchoredZoomOffset, zoomStepScale } from '../interaction/zoom';
import { fitScale, nextZoomStep, prevZoomStep } from '../interaction/zoomable';

export interface ScrollZoomHooks {
  scrollHost(): HTMLDivElement;
  spacer(): HTMLDivElement;
  count(): number;
  zoomMin(): number;
  zoomMax(): number;
  baseScale(): number;
  fitWidthPx(): number;
  fitContentSize(mode: 'width' | 'page'): { width: number; height: number } | null;
  indexAt(contentY: number): number;
  offset(index: number): number;
  height(index: number): number;
  totalHeight(): number;
  recomputeHeights(): void;
  syncSpacerWidth(): void;
  padLeft(): number;
  invalidateRender(): void;
  preview(): void;
  scheduleSettle(): void;
  onScaleChange(scale: number): void;
  relayout(): void;
  mountVisible(): void;
  refitOnResize(): boolean;
}

/** Shared scale state and viewport anchoring. Unit geometry is supplied by the
 * format: DOCX uses variable page offsets and PPTX uses a uniform stride. */
export class ScrollZoomController {
  scale = 1;
  established = false;
  pendingScale: number | null = null;
  pendingAnchor: { x: number; y: number } | null = null;
  prevBase = 0;
  lastFitWidth = 0;
  private _wheelListener: ((event: WheelEvent) => void) | null = null;
  private _resizeObserver: ResizeObserver | null = null;

  constructor(private readonly hooks: ScrollZoomHooks) {}

  bind(container: HTMLElement, host: HTMLDivElement, wheelZoomEnabled: boolean): void {
    if (wheelZoomEnabled) {
      this._wheelListener = (event) => {
        if (!(event.ctrlKey || event.metaKey)) return;
        event.preventDefault();
        if (event.deltaY === 0) return;
        const rect = host.getBoundingClientRect();
        const x = event.clientX - rect.left;
        const y = event.clientY - rect.top;
        this.pendingAnchor = Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
        this.setScale(zoomStepScale(this.scale, event.deltaY, event.deltaMode));
      };
      host.addEventListener('wheel', this._wheelListener as EventListener, { passive: false });
    }
    if (typeof ResizeObserver !== 'undefined') {
      this._resizeObserver = new ResizeObserver(() => this.onResize());
      this._resizeObserver.observe(container);
    }
  }

  /** Apply a pre-load request at first fit, before the opening window paints. */
  establishBase(): boolean {
    const base = this.hooks.baseScale();
    if (base <= 0) return false;
    this.scale = base;
    this.prevBase = base;
    this.lastFitWidth = this.hooks.fitWidthPx();
    this.established = true;
    if (this.pendingScale !== null) {
      const pending = this.pendingScale;
      this.pendingScale = null;
      if (pending !== this.scale) {
        this.scale = pending;
        this.hooks.onScaleChange(pending);
      }
    }
    return true;
  }

  setScale(requested: number): void {
    const next = Math.min(this.hooks.zoomMax(), Math.max(this.effectiveMin(), requested));
    // Consume even for a no-op so a wheel anchor cannot affect a later API call.
    const gestureAnchor = this.pendingAnchor;
    this.pendingAnchor = null;
    if (this.hooks.count() === 0 || !this.established) {
      this.pendingScale = next;
      return;
    }
    if (next === this.scale) return;
    const prevScale = this.scale;
    const anchorY = gestureAnchor?.y ?? 0;
    const host = this.hooks.scrollHost();
    const scrollTop = host.scrollTop;
    const contentY = scrollTop + anchorY;
    const top = this.hooks.indexAt(contentY);
    const oldHeight = this.hooks.height(top);
    const oldOffset = this.hooks.offset(top);
    const fraction = Math.min(1, Math.max(0,
      oldHeight > 0 ? (contentY - oldOffset) / oldHeight : 0,
    ));
    const padLeft = this.hooks.padLeft();
    const scrollLeft = host.scrollLeft || 0;

    // Invalidate old-scale async paints before changing any slot geometry.
    this.hooks.invalidateRender();
    this.scale = next;
    this.hooks.recomputeHeights();
    const totalHeight = this.hooks.totalHeight();
    this.hooks.spacer().style.height = `${totalHeight}px`;
    this.hooks.syncSpacerWidth();

    const maxTop = Math.max(0, totalHeight - host.clientHeight);
    const newContentY = this.hooks.offset(top) + fraction * this.hooks.height(top);
    const reanchoredTop = contentY < oldOffset && top === 0
      ? scrollTop
      : newContentY - anchorY;
    host.scrollTop = Math.min(maxTop, Math.max(0, reanchoredTop));
    if (gestureAnchor) {
      const maxLeft = Math.max(0, (this.hooks.spacer().offsetWidth || 0) - host.clientWidth);
      host.scrollLeft = anchoredZoomOffset(
        scrollLeft, gestureAnchor.x - padLeft, prevScale, next, { maxScroll: maxLeft },
      );
    }
    this.hooks.preview();
    this.hooks.scheduleSettle();
    this.hooks.onScaleChange(next);
  }

  getScale(): number { return this.established ? this.scale : this.pendingScale ?? 1; }
  zoomIn(): void { this.setScale(nextZoomStep(this.getScale())); }
  zoomOut(): void { this.setScale(prevZoomStep(this.getScale(), this.effectiveMin())); }

  effectiveMin(): number {
    const configured = this.hooks.zoomMin();
    return this.established && this.prevBase > 0 ? Math.min(configured, this.prevBase) : configured;
  }

  fit(mode: 'width' | 'page'): void {
    if (this.hooks.count() === 0) return;
    const size = this.hooks.fitContentSize(mode);
    if (!size) return;
    const scale = fitScale({
      contentWidth: size.width,
      contentHeight: size.height,
      containerWidth: this.hooks.fitWidthPx(),
      containerHeight: this.hooks.scrollHost().clientHeight,
    }, mode);
    if (scale > 0) this.setScale(scale);
  }

  onResize(): void {
    // Only a width change alters the fit base, but a height-only resize can
    // expose new units and must still mount them. A zero-width opening layout
    // is retried here. Preserve the user's multiplier over the old base when
    // fitting the new width, then route through setScale so stale renders lose
    // their epoch and the current pixels remain visible until settle.
    if (this.hooks.count() === 0) return;
    if (!this.established) {
      this.hooks.relayout();
      return;
    }
    if (!this.hooks.refitOnResize()) {
      this.lastFitWidth = this.hooks.fitWidthPx();
      this.hooks.mountVisible();
      return;
    }
    const newBase = this.hooks.baseScale();
    if (newBase <= 0) return;
    const newFitWidth = this.hooks.fitWidthPx();
    if (newFitWidth === this.lastFitWidth) {
      this.hooks.mountVisible();
      return;
    }
    this.lastFitWidth = newFitWidth;
    const multiplier = this.prevBase > 0 ? this.scale / this.prevBase : 1;
    this.prevBase = newBase;
    this.setScale(newBase * multiplier);
    // A clamped setScale is a no-op, yet a taller viewport may reveal units.
    this.hooks.mountVisible();
  }

  destroy(): void {
    this.pendingAnchor = null;
    if (this._wheelListener) {
      this.hooks.scrollHost().removeEventListener('wheel', this._wheelListener as EventListener);
      this._wheelListener = null;
    }
    this._resizeObserver?.disconnect();
    this._resizeObserver = null;
  }
}
