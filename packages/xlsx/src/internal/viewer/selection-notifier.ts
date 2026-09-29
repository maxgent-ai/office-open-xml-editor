import type { XlsxSelectionContext, XlsxSelectionState } from '../../selection.js';
import { selectionStatesEqual } from '../../selection.js';

const DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS = 65_536;
const MAX_REENTRANT_SELECTION_NOTIFICATIONS = 100;

/** Selection model, callbacks and scheduling the notifier uses. */
export interface SelectionNotifierHost {
  readonly hostWindow: Window & typeof globalThis;
  isDestroyed(): boolean;
  selectionState(): XlsxSelectionState | null;
  onSelectionStateChange(): ((selection: XlsxSelectionState | null) => void) | undefined;
  onSelectionContextChange(): ((context: XlsxSelectionContext | null) => void) | undefined;
  /** Bounded context snapshot for a notification. */
  readContext(maxTextCharacters: number): XlsxSelectionContext | null;
  /** Re-enter the engine's selection-change seam from a queued delivery. */
  emitSelectionChange(): void;
  /** The engine's context-notification seam (routes back to this notifier). */
  scheduleSelectionContextNotification(): void;
}

/**
 * Delivers `onSelectionStateChange` and frame-coalesced
 * `onSelectionContextChange` notifications. Re-entrant changes made from a
 * callback are coalesced into a bounded microtask chain so a callback feedback
 * cycle cannot monopolize the main thread.
 */
export class SelectionNotifier {
  private lastNotified: XlsxSelectionState | null = null;
  private emitting = false;
  private pending = false;
  private scheduled = false;
  private chainLength = 0;
  private contextFrame: number | null = null;
  private contextMicrotask = false;

  constructor(private readonly host: SelectionNotifierHost) {}

  /** Queue one onSelectionContextChange delivery for the next frame. */
  scheduleContextNotification(): void {
    if (!this.host.onSelectionContextChange() || this.host.isDestroyed() ||
        this.contextFrame !== null ||
        this.contextMicrotask) return;
    const notify = () => {
      this.contextFrame = null;
      this.contextMicrotask = false;
      if (this.host.isDestroyed()) return;
      const context = this.host.readContext(DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS);
      this.host.onSelectionContextChange()?.(context ? structuredClone(context) : null);
    };
    if (typeof this.host.hostWindow.requestAnimationFrame === 'function') {
      this.contextFrame = this.host.hostWindow.requestAnimationFrame(notify);
    } else {
      this.contextMicrotask = true;
      queueMicrotask(notify);
    }
  }

  /** Deliver onSelectionStateChange for the current selection, coalescing
   *  re-entrant changes made by the callback itself. */
  emit(): void {
    const state = this.host.selectionState();
    if (!selectionStatesEqual(state, this.lastNotified)) {
      this.host.scheduleSelectionContextNotification();
    }
    if (this.emitting) {
      this.pending = true;
      this.scheduleNotification();
      return;
    }
    this.pending = false;
    if (selectionStatesEqual(state, this.lastNotified)) {
      this.finishChain();
      return;
    }

    if (this.chainLength >= MAX_REENTRANT_SELECTION_NOTIFICATIONS) {
      // A callback feedback cycle must not monopolize the main thread. The
      // canonical state remains authoritative; only notifications beyond the
      // documented per-chain safety limit are suppressed.
      this.lastNotified = state ? structuredClone(state) : null;
      this.finishChain();
      return;
    }
    this.chainLength++;
    this.lastNotified = state ? structuredClone(state) : null;
    this.emitting = true;
    try {
      this.host.onSelectionStateChange()?.(state ? structuredClone(state) : null);
    } finally {
      this.emitting = false;
      if (this.pending ||
          !selectionStatesEqual(this.host.selectionState(), this.lastNotified)) {
        this.scheduleNotification();
      } else {
        this.finishChain();
      }
    }
  }

  private scheduleNotification(): void {
    if (this.scheduled || this.host.isDestroyed()) return;
    this.scheduled = true;
    queueMicrotask(() => {
      this.scheduled = false;
      if (!this.host.isDestroyed()) this.host.emitSelectionChange();
    });
  }

  private finishChain(): void {
    this.pending = false;
    this.chainLength = 0;
  }

  /** Teardown: cancel a queued context delivery and reset the chain. */
  destroy(): void {
    if (this.contextFrame !== null) {
      this.host.hostWindow.cancelAnimationFrame(this.contextFrame);
      this.contextFrame = null;
    }
    this.contextMicrotask = false;
    this.lastNotified = null;
    this.finishChain();
  }
}
