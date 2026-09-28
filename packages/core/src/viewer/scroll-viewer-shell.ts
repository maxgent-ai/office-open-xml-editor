import { eventTargetsDataAttributeWithin } from '../internal/dom-interaction-boundary';

/** Owns the scroll surface and its DOM listeners. Format adapters supply the
 * load/visibility policy and the action that clears an active comment. */
export class ScrollViewerShell {
  readonly wrapper: HTMLDivElement;
  readonly scrollHost: HTMLDivElement;
  readonly spacer: HTMLDivElement;
  private readonly onScroll: () => void;
  private readonly onPointerDown: ((event: PointerEvent) => void) | null;

  constructor(
    container: HTMLElement,
    options: {
      background?: string;
      comments: boolean;
      onScroll: () => void;
      onOutsideComment: () => void;
    },
  ) {
    const doc = container.ownerDocument ?? document;
    this.wrapper = doc.createElement('div');
    this.wrapper.style.cssText = 'position:relative;width:100%;height:100%;overflow:hidden;';
    this.scrollHost = doc.createElement('div');
    this.scrollHost.style.cssText = 'position:absolute;inset:0;overflow:auto;';
    this.scrollHost.style.scrollbarGutter = 'stable';
    if (options.background) this.scrollHost.style.background = options.background;
    this.spacer = doc.createElement('div');
    this.spacer.style.cssText = 'position:absolute;top:0;left:0;width:1px;height:0;pointer-events:none;';
    this.scrollHost.appendChild(this.spacer);
    this.wrapper.appendChild(this.scrollHost);
    container.appendChild(this.wrapper);

    this.onScroll = options.onScroll;
    this.scrollHost.addEventListener('scroll', this.onScroll);
    this.onPointerDown = options.comments ? (event) => {
      if (!eventTargetsDataAttributeWithin(event, this.wrapper, 'ooxmlCommentId')) {
        options.onOutsideComment();
      }
    } : null;
    if (this.onPointerDown) doc.addEventListener('pointerdown', this.onPointerDown);
  }

  destroy(): void {
    this.scrollHost.removeEventListener('scroll', this.onScroll);
    if (this.onPointerDown) {
      this.wrapper.ownerDocument.removeEventListener('pointerdown', this.onPointerDown);
    }
    this.wrapper.remove();
  }
}
