import { resolveItemStartScrollTop } from '../layout/virtual-scroll';

export interface ScrollNavigationHooks {
  host(): HTMLDivElement;
  spacer(): HTMLDivElement;
  count(): number;
  established(): boolean;
  offset(unit: number): number;
  height(unit: number): number;
  indexAt(contentY: number): number;
  totalHeight(): number;
  width(unit: number): number;
  padLeft(): number;
  marginOrigin(): number;
  mount(): void;
}

/** Scrolls authored units and points into view using current virtual geometry. */
export class ScrollNavigationController {
  constructor(private readonly hooks: ScrollNavigationHooks) {}

  scrollToUnit(index: number, options?: { behavior?: 'auto' | 'smooth' }): void {
    if (this.hooks.count() === 0 || !this.hooks.established()) return;
    const unit = Math.max(0, Math.min(index, this.hooks.count() - 1));
    const host = this.hooks.host();
    const maxTop = Math.max(0, this.hooks.totalHeight() - host.clientHeight);
    const top = resolveItemStartScrollTop(this.hooks.offset(unit), maxTop);
    if (typeof host.scrollTo === 'function') {
      host.scrollTo({ top, behavior: options?.behavior ?? 'auto' });
    } else {
      host.scrollTop = top;
    }
    this.hooks.mount();
  }

  /** Point coordinates are CSS px relative to the unit's authored left edge. */
  scrollToPoint(unit: number, x: number, y: number, options?: { behavior?: 'auto' | 'smooth' }): void {
    const host = this.hooks.host();
    const leftEdge = Math.max(this.hooks.padLeft(),
      (host.clientWidth - this.hooks.width(unit)) / 2) + this.hooks.marginOrigin();
    const maxTop = Math.max(0, this.hooks.totalHeight() - host.clientHeight);
    const spacer = this.hooks.spacer();
    const spacerWidth = spacer.offsetWidth || Number.parseFloat(spacer.style.width) || 0;
    const maxLeft = Math.max(0, spacerWidth - host.clientWidth);
    const top = Math.min(maxTop, Math.max(0,
      this.hooks.offset(unit) + y - host.clientHeight / 2));
    const left = Math.min(maxLeft, Math.max(0,
      leftEdge + x - host.clientWidth / 2));
    if (typeof host.scrollTo === 'function') {
      host.scrollTo({ top, left, behavior: options?.behavior ?? 'auto' });
    } else {
      host.scrollTop = top;
      host.scrollLeft = left;
    }
    this.hooks.mount();
  }

  contentAtViewportY(y: number): { unit: number; frac: number } {
    const contentY = this.hooks.host().scrollTop + y;
    const unit = this.hooks.indexAt(contentY);
    const height = this.hooks.height(unit);
    const frac = height > 0 ? Math.min(1, Math.max(0,
      (contentY - this.hooks.offset(unit)) / height)) : 0;
    return { unit, frac };
  }

  viewportYOf(unit: number, frac: number): number {
    return this.hooks.offset(unit) + frac * this.hooks.height(unit) - this.hooks.host().scrollTop;
  }
}
