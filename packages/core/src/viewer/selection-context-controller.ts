import { renderCanvasElementOutline, type CanvasElementOutline } from '../internal/canvas-viewer-mechanics';

export interface SelectionSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  elementLayer: HTMLDivElement | null;
}

export interface SelectionContextHooks<Context extends { kind: string }, ElementContext, Resource, Slot extends SelectionSlot> {
  wrapper(): HTMLDivElement;
  scrollHost(): HTMLDivElement;
  slots(): ReadonlyMap<number, Slot>;
  resource(): Resource | null;
  destroyed(): boolean;
  textSelectionEnabled(): boolean;
  elementSelectionEnabled(): boolean;
  textSelected(): boolean;
  getContext(): Context | null;
  hitTest(resource: Resource, unit: number, xRatio: number, yRatio: number, canvasWidth: number): Promise<ElementContext | null>;
  outline(resource: Resource, unit: number, context: ElementContext): CanvasElementOutline | null;
  onChange(context: Context | null): void;
  onContextMenu?(event: MouseEvent, getContext: () => Promise<Context | null>): void;
  reportError(error: unknown): void;
}

/** Selection event and outline lifecycle. Formats provide their text locators,
 * element hit-test coordinates, and normalized outline geometry. */
export class SelectionContextController<Context extends { kind: string }, ElementContext, Resource, Slot extends SelectionSlot> {
  elementContext: ElementContext | null = null;
  private _generation = 0;
  private _contextKey = 'null';
  private _selectionChangeListener: (() => void) | null = null;
  private _elementClickListener: ((event: MouseEvent) => void) | null = null;
  private _contextMenuListener: ((event: MouseEvent) => void) | null = null;

  constructor(private readonly hooks: SelectionContextHooks<Context, ElementContext, Resource, Slot>) {}

  bind(hasChangeListener: boolean, hasContextMenu: boolean): void {
    const wrapper = this.hooks.wrapper();
    const host = this.hooks.scrollHost();
    if (this.hooks.textSelectionEnabled() && (hasChangeListener || this.hooks.elementSelectionEnabled())) {
      this._selectionChangeListener = () => this.emitChange();
      wrapper.ownerDocument.addEventListener('selectionchange', this._selectionChangeListener);
    }
    if (this.hooks.elementSelectionEnabled()) {
      this._elementClickListener = (event) => {
        void this.onElementClick(event).catch((error) => this.hooks.reportError(error));
      };
      host.addEventListener('click', this._elementClickListener);
    }
    if (hasContextMenu && this.hooks.onContextMenu) {
      this._contextMenuListener = (event) => {
        let context: Promise<Context | null> | undefined;
        this.hooks.onContextMenu?.(event, () => context ??= this.resolveContextAt(event));
      };
      host.addEventListener('contextmenu', this._contextMenuListener);
    }
  }

  emitChange(): void {
    const context = this.hooks.getContext();
    if (context?.kind === 'text') {
      this._generation++;
      this.elementContext = null;
      this.redrawOutlines();
    }
    const key = JSON.stringify(context);
    if (key === this._contextKey) return;
    this._contextKey = key;
    this.hooks.onChange(context ? structuredClone(context) : null);
  }

  setElementContext(context: ElementContext | null): void {
    this.elementContext = context ? structuredClone(context) : null;
    this.redrawOutlines();
    this.emitChange();
  }

  invalidateElementContext(notify = true): void {
    this._generation++;
    this.elementContext = null;
    this.redrawOutlines();
    if (notify) this.emitChange();
  }

  clearElementContext(): void { this.elementContext = null; }

  redrawOutlines(): void {
    for (const [unit, slot] of this.hooks.slots()) this.redrawOutlineForSlot(unit, slot);
  }

  redrawOutlineForSlot(unit: number, slot: Slot): void {
    const context = this.elementContext;
    const resource = this.hooks.resource();
    renderCanvasElementOutline(slot.elementLayer,
      context && resource ? this.hooks.outline(resource, unit, context) : null);
  }

  private async onElementClick(event: MouseEvent): Promise<void> {
    if (this.hooks.destroyed() || event.defaultPrevented || event.button !== 0) return;
    await this.resolveContextAt(event);
  }

  async resolveContextAt(event: MouseEvent): Promise<Context | null> {
    const resource = this.hooks.resource();
    if (this.hooks.destroyed() || !resource) return null;
    if (this.hooks.textSelectionEnabled() && this.hooks.textSelected()) {
      this.emitChange();
      return this.hooks.destroyed() ? null : this.hooks.getContext();
    }
    if (!this.hooks.elementSelectionEnabled()) return this.hooks.getContext();
    const target = event.target as Node | null;
    const entry = [...this.hooks.slots()].find(([, slot]) => target !== null && slot.wrapper.contains(target));
    if (!entry) {
      this.invalidateElementContext();
      return null;
    }
    const [unit, slot] = entry;
    const rect = slot.canvas.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      this.invalidateElementContext();
      return null;
    }
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
      this.invalidateElementContext();
      return null;
    }
    const generation = ++this._generation;
    let context: ElementContext | null;
    try {
      context = await this.hooks.hitTest(resource, unit, x / rect.width, y / rect.height, rect.width);
    } catch (error) {
      if (this.hooks.destroyed() || generation !== this._generation || resource !== this.hooks.resource()) return null;
      throw error;
    }
    if (this.hooks.destroyed() || generation !== this._generation || resource !== this.hooks.resource()) return null;
    this.setElementContext(context);
    return this.hooks.destroyed() ? null : this.hooks.getContext();
  }

  destroy(): void {
    this._generation++;
    const wrapper = this.hooks.wrapper();
    const host = this.hooks.scrollHost();
    if (this._selectionChangeListener) {
      wrapper.ownerDocument.removeEventListener('selectionchange', this._selectionChangeListener);
      this._selectionChangeListener = null;
    }
    if (this._elementClickListener) {
      host.removeEventListener('click', this._elementClickListener);
      this._elementClickListener = null;
    }
    if (this._contextMenuListener) {
      host.removeEventListener('contextmenu', this._contextMenuListener);
      this._contextMenuListener = null;
    }
    this.elementContext = null;
  }
}
