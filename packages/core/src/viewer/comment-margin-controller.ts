export interface CommentMarginHooks {
  container(): HTMLElement;
  scrollHost(): HTMLDivElement;
  spacer(): HTMLDivElement;
  enabled(): boolean;
  cards(): boolean;
  hasDisplayableComments(): boolean;
  requestedSide(): 'auto' | 'left' | 'right' | undefined;
  zoom(): number;
  gapPx: number;
  widthPx: number;
  fontSizePx: number;
}

/** Native horizontal review-rail geometry shared by the scroll viewers.
 * Each format decides when authored comments become displayable. */
export class CommentMarginController {
  private _originPx = 0;

  constructor(private readonly hooks: CommentMarginHooks) {}

  enabled(): boolean { return this.hooks.enabled(); }
  zoom(): number { return this.hooks.zoom(); }
  get originPx(): number { return this._originPx; }

  hasMargin(): boolean {
    return this.hooks.enabled() && this.hooks.cards() && this.hooks.hasDisplayableComments();
  }

  side(): 'left' | 'right' {
    const requested = this.hooks.requestedSide();
    if (requested === 'left' || requested === 'right') return requested;
    const container = this.hooks.container();
    const computedDirection = container.ownerDocument.defaultView?.getComputedStyle?.(container).direction;
    const direction = computedDirection || container.dir || container.style.direction;
    return direction === 'rtl' ? 'left' : 'right';
  }

  extent(): number {
    return this.hasMargin() ? (this.hooks.gapPx + this.hooks.widthPx) * this.hooks.zoom() : 0;
  }

  syncMargin(margin: HTMLDivElement | null): void {
    if (!margin) return;
    margin.style.display = this.hasMargin() ? '' : 'none';
    const zoom = this.hooks.zoom();
    const offset = `calc(100% + ${this.hooks.gapPx * zoom}px)`;
    margin.style.left = this.side() === 'right' ? offset : '';
    margin.style.right = this.side() === 'left' ? offset : '';
    margin.style.width = `${this.hooks.widthPx * zoom}px`;
    margin.style.fontSize = `${this.hooks.fontSizePx}px`;
    margin.dataset.ooxmlCommentZoom = String(zoom);
  }

  syncSpacerWidth(contentWidth: number, left: number, right: number): void {
    const marginExtent = this.extent();
    const next = this.side() === 'left' ? marginExtent : 0;
    const delta = next - this._originPx;
    const scrollHost = this.hooks.scrollHost();
    const targetScrollLeft = Math.max(0, scrollHost.scrollLeft + delta);
    // The browser clamps scrollLeft to the current range at assignment time.
    this.hooks.spacer().style.width = `${contentWidth + marginExtent + left + right}px`;
    if (delta === 0) return;
    this._originPx = next;
    (scrollHost.style as CSSStyleDeclaration & Record<string, string>)[
      '--ooxml-review-origin-x'
    ] = `${next}px`;
    scrollHost.scrollLeft = targetScrollLeft;
  }

  destroy(): void {
    this._originPx = 0;
  }
}
