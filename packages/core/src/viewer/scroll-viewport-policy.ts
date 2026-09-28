export interface ScrollViewportOptions {
  width?: number;
  gap?: number;
  overscan?: number;
  paddingTop?: number;
  paddingBottom?: number;
  paddingLeft?: number;
  paddingRight?: number;
  dpr?: number;
}

/** Default canvas-only shadow; a false option disables it without moving layout. */
export const DEFAULT_SCROLL_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';

/** Format-neutral desk gutters, fit width, overscan and device pixel ratio. */
export class ScrollViewportPolicy {
  constructor(private readonly hooks: {
    options(): ScrollViewportOptions;
    container(): HTMLElement;
    scrollHost(): HTMLDivElement;
  }) {}

  gap(): number { return this.hooks.options().gap ?? 16; }
  overscan(): number { return this.hooks.options().overscan ?? 1; }

  verticalPadding(): { leading: number; trailing: number } {
    const options = this.hooks.options();
    const gap = this.gap();
    return { leading: options.paddingTop ?? gap, trailing: options.paddingBottom ?? gap };
  }

  horizontalPadding(): { left: number; right: number } {
    const options = this.hooks.options();
    const gap = this.gap();
    return { left: options.paddingLeft ?? gap, right: options.paddingRight ?? gap };
  }

  /** Explicit width is the authored unit width; only inferred fit subtracts gutters. */
  fitWidth(): number {
    const explicit = this.hooks.options().width;
    if (explicit && explicit > 0) return explicit;
    const width = this.hooks.scrollHost().clientWidth || this.hooks.container().clientWidth;
    if (width <= 0) return 0;
    const { left, right } = this.horizontalPadding();
    return Math.max(0, width - left - right);
  }

  dpr(): number {
    return this.hooks.options().dpr ??
      (typeof window !== 'undefined' ? window.devicePixelRatio || 1 : 1);
  }
}
