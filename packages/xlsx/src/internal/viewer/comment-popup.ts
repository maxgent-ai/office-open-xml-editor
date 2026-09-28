import type { XlsxComment } from '../../types.js';
import type { CellAddress } from '../../selection.js';
import type { ReadOnlyCommentThread } from '@silurus/ooxml-core/internal/read-only-comment-contract';
import { parseA1 } from '../../a1.js';
import { computeCommentPopupPosition } from '../../comment-popup.js';
import type { SheetOverlayHost } from '../sheet-surface.js';

type XlsxCommentUiRuntime = typeof import('../../comment-ui-runtime.js');
let xlsxCommentUiRuntimePromise: Promise<XlsxCommentUiRuntime> | undefined;

function loadXlsxCommentUiRuntime(): Promise<XlsxCommentUiRuntime> {
  return xlsxCommentUiRuntimePromise ??= import('../../comment-ui-runtime.js');
}

/** Delay (ms) before a hovered comment popup appears. A short hover dwell
 *  prevents the popup from flickering while the cursor sweeps across many
 *  commented cells; ~150ms is the common tooltip-show threshold (responsive yet
 *  long enough to suppress transient passes). Excel itself uses a comparable
 *  short hover delay before showing a note. */
const COMMENT_POPUP_DELAY_MS = 150;
/** Max width of the comment popup body (CSS px). */
export const COMMENT_POPUP_MAX_W = 280;
/** Max height before the body scrolls/clips (CSS px). */
export const COMMENT_POPUP_MAX_H = 200;

type CellRect = { x: number; y: number; w: number; h: number };

/** Build a `"row:col"` → comment index. Parses each `XlsxComment.cellRef` with
 *  the shared {@link parseA1}; later refs win on a collision (Excel allows at
 *  most one note per cell, so this is moot in practice). */
export function createCommentMap(comments: readonly XlsxComment[]): Map<string, XlsxComment> {
  const map = new Map<string, XlsxComment>();
  for (const c of comments) {
    const p = parseA1(c.cellRef);
    if (p) map.set(`${p.row}:${p.col}`, c);
  }
  return map;
}

/** Viewer state the cell-anchored comment popup reads from its engine. */
export interface CommentPopupHost {
  readonly ownerDocument: Document;
  readonly canvasArea: HTMLDivElement;
  readonly overlayHost: SheetOverlayHost;
  currentSheet(): number;
  isRtl(): boolean;
  isDestroyed(): boolean;
  cellRect(row: number, col: number): CellRect | null;
  screenX(logicalX: number, width: number): number;
  reportError(error: unknown): void;
}

/**
 * Excel-style note popup for the current sheet's comments: hover dwell,
 * immediate display for touch/keyboard, the lazily loaded shared comment-card
 * UI, the screen-reader announcement and cell-anchored positioning. Owns its
 * show timer, position frame and the ResizeObserver that re-anchors it.
 */
export class CommentPopup {
  /** DOM overlay element that shows the hovered cell's comment. */
  readonly popup: HTMLDivElement;
  /** `"row:col"` → comment for the displayed sheet, rebuilt on every show. */
  commentMap = new Map<string, XlsxComment>();
  /** `"row:col"` of the cell whose popup is currently shown (or pending), so a
   *  pointermove within the same cell doesn't restart the show timer. */
  private popupKey: string | null = null;
  /** Pending show timer (see {@link COMMENT_POPUP_DELAY_MS}). */
  private timer: ReturnType<typeof setTimeout> | null = null;
  private cell: CellAddress | null = null;
  private positionScheduled = false;
  private resizeObserver: ResizeObserver | null = null;
  private ui: XlsxCommentUiRuntime | null = null;
  private renderGeneration = 0;

  constructor(private readonly host: CommentPopupHost) {
    this.popup = host.overlayHost.comment;
    const ResizeObserverClass = host.ownerDocument.defaultView?.ResizeObserver ??
      globalThis.ResizeObserver;
    if (ResizeObserverClass) {
      this.resizeObserver = new ResizeObserverClass(() => {
        this.schedulePosition();
      });
      this.resizeObserver.observe(this.popup);
    }
  }

  /** Index the displayed sheet's (visibility-filtered) comments. */
  setComments(comments: readonly XlsxComment[]): void {
    this.commentMap = createCommentMap(comments);
  }

  /** The comment on `cell`, if any. */
  commentAt(cell: CellAddress): XlsxComment | undefined {
    return this.commentMap.get(`${cell.row}:${cell.col}`);
  }

  /** Whether the popup is currently displayed. */
  isOpen(): boolean {
    return this.popup.style.display !== 'none';
  }

  /** Whether `node` lies inside the popup (pointer moved onto it). */
  contains(node: Node): boolean {
    return this.popup.contains(node);
  }

  /** Show the popup for the comment on `cell` after the hover dwell, anchored to
   *  the cell's current on-screen rect. No-op when the cell carries no comment.
   *  Re-hovering the same cell does not restart the timer. */
  scheduleForCell(cell: CellAddress): void {
    const key = `${cell.row}:${cell.col}`;
    const comment = this.commentMap.get(key);
    if (!comment) {
      this.hide();
      return;
    }
    if (this.popupKey === key) return; // already shown / pending here
    this.hide();
    this.popupKey = key;
    this.timer = setTimeout(() => {
      this.timer = null;
      void this.show(cell, comment).catch((error) => this.host.reportError(error));
    }, COMMENT_POPUP_DELAY_MS);
  }

  /** Load (once) the shared comment-card runtime. */
  async loadUi(): Promise<XlsxCommentUiRuntime> {
    const commentUi = this.ui ?? await loadXlsxCommentUiRuntime();
    if (!this.host.isDestroyed()) this.ui = commentUi;
    return commentUi;
  }

  /** Immediately render the popup for `comment` anchored to `cell` (used by the
   *  hover-dwell timer and by touch selection, which has no hover). */
  async show(cell: CellAddress, comment: XlsxComment): Promise<void> {
    if (!this.host.cellRect(cell.row, cell.col)) return;
    const generation = ++this.renderGeneration;
    const commentUi = await this.loadUi();
    if (this.host.isDestroyed() || generation !== this.renderGeneration) return;
    if (!this.host.cellRect(cell.row, cell.col)) return;
    this.cell = cell;

    // Use the same card structure and default theme as the DOCX/PPTX margins;
    // XLSX owns only the cell-anchored popup geometry.
    const occurrenceKey = `sheet:${this.host.currentSheet()}:cell:${comment.cellRef}:comment:${comment.id ?? 'root'}`;
    const thread: ReadOnlyCommentThread = {
      occurrenceKey,
      root: {
        messageKey: `${occurrenceKey}:root`,
        sourceId: comment.id,
        author: comment.author,
        date: comment.date,
        text: comment.rootText ?? comment.text,
        status: comment.resolved ? 'resolved' : 'active',
      },
      replies: (comment.replies ?? []).map((reply, index) => ({
        messageKey: `${occurrenceKey}:reply:${reply.id ?? index}`,
        sourceId: reply.id,
        author: reply.author,
        date: reply.date,
        text: reply.text,
        status: reply.resolved ? 'resolved' : 'active',
      })),
    };
    commentUi.paintReadOnlyCommentCard(this.popup, thread, {
      interactive: false,
      standalone: true,
    });
    const rootText = (comment.rootText ?? comment.text).trim();
    const byAuthor = comment.author?.trim() ? ` by ${comment.author.trim()}` : '';
    const replyCount = comment.replies?.length ?? 0;
    const replies = replyCount === 0
      ? ''
      : `; ${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}`;
    this.host.overlayHost.announceComment(
      `Comment on ${comment.cellRef}${byAuthor}${rootText ? `: ${rootText}` : ''}${replies}`,
    );
    this.popup.dataset.ooxmlCommentUi = 'popup';
    this.popup.style.maxWidth = `${COMMENT_POPUP_MAX_W}px`;
    this.popup.style.maxHeight = `${COMMENT_POPUP_MAX_H}px`;

    // Anchor to the cell's *screen* rect (RTL already mirrored by screenX), then
    // run the pure position calc against the popup's measured size. Make it
    // visible (off-screen) first so offsetWidth/Height reflect the wrapped text.
    this.popup.style.left = '-9999px';
    this.popup.style.top = '-9999px';
    this.popup.style.display = '';
    this.position();
  }

  private schedulePosition(): void {
    if (this.positionScheduled || !this.cell) return;
    this.positionScheduled = true;
    const position = (): void => {
      this.positionScheduled = false;
      this.position();
    };
    const ownerWindow = this.host.ownerDocument.defaultView;
    if (ownerWindow?.requestAnimationFrame) ownerWindow.requestAnimationFrame(position);
    else queueMicrotask(position);
  }

  private position(): void {
    const cell = this.cell;
    if (!cell || this.popup.style.display === 'none') return;
    const rect = this.host.cellRect(cell.row, cell.col);
    if (!rect) return;
    const screenLeft = this.host.screenX(rect.x, rect.w);
    const pos = computeCommentPopupPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      popup: { w: this.popup.offsetWidth, h: this.popup.offsetHeight },
      viewport: { w: this.host.canvasArea.clientWidth, h: this.host.canvasArea.clientHeight },
      rtl: this.host.isRtl(),
    });
    this.host.overlayHost.showComment(pos.left, pos.top);
  }

  /** Hide the popup and cancel any pending show. Called on cell-out, scroll,
   *  sheet switch and destroy. */
  hide(): void {
    this.renderGeneration++;
    if (this.timer !== null) {
      clearTimeout(this.timer);
      this.timer = null;
    }
    this.popupKey = null;
    this.cell = null;
    this.host.overlayHost.hideComment();
    this.popup.replaceChildren();
  }

  /** Teardown: stop observing, cancel pending work and drop the index. */
  destroy(): void {
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    this.hide();
    this.commentMap.clear();
  }
}
