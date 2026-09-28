import type { ViewerCommentConnectorOptions } from '../comment-ui';
import type { ReadOnlyCommentMarginGeometry } from '../internal/read-only-comment-decoration';

export interface CommentOverlaySlot {
  commentDecorationLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentGeometry: ReadOnlyCommentMarginGeometry | null;
  renderedScale: number;
}

type DecorationRuntime = Pick<typeof import('../internal/read-only-comment-decoration'),
  'buildReadOnlyCommentDecoration' | 'projectReadOnlyCommentMarginScroll'>;

export interface CommentOverlayHooks<Slot extends CommentOverlaySlot> {
  slots(): ReadonlyMap<number, Slot>;
  scale(): number;
  destroyed(): boolean;
  ownerWindow(): Window | null;
  width(unit: number): number;
  height(unit: number): number;
  side(): 'left' | 'right';
  marginExtent(): number;
  connectorOptions(): ViewerCommentConnectorOptions | undefined;
  runtime(): DecorationRuntime | null;
  redrawComments(unit: number, slot: Slot): void;
}

/** Built-in connector projection and one-frame comment geometry scheduling. */
export class CommentOverlayController<Slot extends CommentOverlaySlot> {
  private _scheduled = false;
  private _frame: number | null = null;
  private readonly _pending = new Map<number, { slot: Slot; connectorsOnly: boolean }>();

  constructor(private readonly hooks: CommentOverlayHooks<Slot>) {}

  drawConnectors(unit: number, slot: Slot): void {
    const layer = slot.commentDecorationLayer;
    const margin = slot.commentMargin;
    const geometry = slot.commentGeometry;
    const connectorOptions = this.hooks.connectorOptions();
    const runtime = this.hooks.runtime();
    if (!layer || !margin || !geometry || !connectorOptions || !runtime) return;
    const width = this.hooks.width(unit);
    const height = this.hooks.height(unit);
    const side = this.hooks.side();
    const marginExtent = this.hooks.marginExtent();
    runtime.buildReadOnlyCommentDecoration(
      layer,
      Object.freeze({
        surfaceBounds: Object.freeze({
          x: side === 'left' ? -marginExtent : 0,
          y: 0,
          width: width + marginExtent,
          height,
        }),
        contentBounds: Object.freeze({ x: 0, y: 0, width, height }),
        side,
        threads: runtime.projectReadOnlyCommentMarginScroll(geometry, margin.scrollTop),
      }),
      {
        route: connectorOptions.route ?? 'bezier',
        stroke: connectorOptions.stroke ?? 'solid',
        color: connectorOptions.color,
        activeColor: connectorOptions.activeColor,
      },
    );
  }

  /** A full card measurement dominates a connector-only refresh for a slot. */
  schedule(unit: number, slot: Slot, connectorsOnly = false): void {
    const current = this._pending.get(unit);
    this._pending.set(unit, {
      slot,
      connectorsOnly: current?.slot === slot
        ? current.connectorsOnly && connectorsOnly
        : connectorsOnly,
    });
    if (this._scheduled) return;
    this._scheduled = true;
    const flush = (): void => {
      this._scheduled = false;
      this._frame = null;
      const pending = [...this._pending];
      this._pending.clear();
      if (this.hooks.destroyed()) return;
      for (const [pendingUnit, entry] of pending) {
        if (this.hooks.slots().get(pendingUnit) !== entry.slot ||
            entry.slot.renderedScale !== this.hooks.scale()) continue;
        if (entry.connectorsOnly) this.drawConnectors(pendingUnit, entry.slot);
        else this.hooks.redrawComments(pendingUnit, entry.slot);
      }
    };
    const ownerWindow = this.hooks.ownerWindow();
    if (ownerWindow?.requestAnimationFrame) this._frame = ownerWindow.requestAnimationFrame(flush);
    else queueMicrotask(flush);
  }

  destroy(): void {
    if (this._frame !== null) this.hooks.ownerWindow()?.cancelAnimationFrame?.(this._frame);
    this._frame = null;
    this._scheduled = false;
    this._pending.clear();
  }
}
