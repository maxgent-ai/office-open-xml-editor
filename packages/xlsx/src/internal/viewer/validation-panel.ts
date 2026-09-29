import type { Worksheet } from '../../types.js';
import type { XlsxWorkbook } from '../../workbook.js';
import type { CellAddress } from '../../selection.js';
import { HEADER_W, HEADER_H } from '../../renderer.js';
import { findListValidationAt } from '../../data-validation.js';
import { computeValidationPanelPosition, type ResolvedList } from '../../validation-list.js';
import type { CanvasSurface, SheetOverlayHost } from '../sheet-surface.js';
import type { SheetSelectionMode } from '../sheet-viewer-runtime.js';
import { ListenerScope } from './listener-scope.js';

type CellRect = { x: number; y: number; w: number; h: number };

/** Viewer state the list-validation dropdown reads from its engine. */
export interface ValidationPanelHost {
  readonly ownerDocument: Document;
  readonly canvasArea: HTMLDivElement;
  readonly surface: CanvasSurface;
  readonly overlayHost: SheetOverlayHost;
  worksheet(): Worksheet | null;
  workbook(): XlsxWorkbook | null;
  currentSheet(): number;
  activeCell(): CellAddress | null;
  selectionMode(): SheetSelectionMode;
  scale(): number;
  isRtl(): boolean;
  isDestroyed(): boolean;
  cellRect(row: number, col: number): CellRect | null;
  screenX(logicalX: number, width: number): number;
}

/**
 * The list data-validation dropdown (ECMA-376 §18.3.1.33): the in-cell arrow
 * button drawn on the active cell and the display-only panel listing the
 * allowed values. The panel is read-only — hovering highlights an item but
 * picking a value never changes the cell. Owns the outside-click document
 * listener and the per-item hover listeners.
 */
export class ValidationPanel {
  /** DOM overlay listing a list-validated cell's allowed values. Lives in
   *  canvasArea above the scrollHost; unlike the comment popup this is a click
   *  target (`pointer-events:auto`). */
  readonly panel: HTMLDivElement;
  /** `"row:col"` of the cell whose panel is pending or open, or null. Claiming
   *  the key before async range resolution lets a re-click cancel the request. */
  panelKey: string | null = null;
  private requestGeneration = 0;
  /** Screen rect (canvasArea CSS px) of the dropdown arrow button last drawn by
   *  {@link drawDropdown}, so pointerdown can hit-test it. Null when no arrow is
   *  currently visible. */
  private arrowRect: CellRect | null = null;
  /** Document-level pointerdown listener that closes the panel on an outside
   *  click; installed only while the panel is open. */
  private outsideHandler: ((e: PointerEvent) => void) | null = null;
  /** Hover listeners on the currently rendered value items. */
  private itemListeners = new ListenerScope();

  constructor(private readonly host: ValidationPanelHost) {
    this.panel = host.overlayHost.validation;
  }

  /** Whether the value panel is currently shown. */
  isOpen(): boolean {
    return this.panel.style.display !== 'none';
  }

  /** Whether a client point lands on the dropdown arrow drawn for the active
   *  cell. The arrow rect is in canvasArea space, so map the client point
   *  through canvasArea's box. */
  hitsArrow(clientX: number, clientY: number): boolean {
    const ar = this.arrowRect;
    if (!ar) return false;
    const { x: ax, y: ay } = this.host.surface.localPoint(clientX, clientY);
    return ax >= ar.x && ax <= ar.x + ar.w && ay >= ar.y && ay <= ar.y + ar.h;
  }

  /** Draw the Excel list-validation dropdown button just outside the
   *  bottom-right corner of the *active* cell when that cell is covered by a
   *  `list` data-validation rule. Anchored to the single active cell (not the
   *  whole range) to mirror Excel, which attaches the button to the active
   *  cell of the selection. */
  drawDropdown(): void {
    // The overlay is rebuilt on every selection / scroll change, so the
    // arrow's hit-test rect is recomputed here each time (cleared when no arrow
    // is currently shown).
    this.arrowRect = null;
    if (this.host.selectionMode() !== 'cells') return;
    const ws = this.host.worksheet();
    const active = this.host.activeCell();
    if (!ws || !active) return;
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;

    const rect = this.host.cellRect(active.row, active.col);
    if (!rect) return;

    // Excel's dropdown button is a fixed square sized to the cell height,
    // clamped to a sensible range so it stays usable at small zoom and doesn't
    // dominate tall rows. The arrow glyph is centered inside.
    const cs = this.host.scale();
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const side = Math.max(14, Math.min(rect.h, 22 * cs));
    // Button sits flush to the right of the cell, top-aligned with it.
    const btnLogicalX = rect.x + rect.w;
    const btnY = rect.y;
    // Cull when the active cell (hence its button) is scrolled behind the
    // fixed headers.
    if (btnLogicalX + side <= headerW || btnY + side <= headerH) return;

    const screenLeft = this.host.screenX(btnLogicalX, side);

    const btn = this.host.ownerDocument.createElement('div');
    btn.setAttribute('data-xlsx-validation-dropdown', '');
    btn.style.cssText =
      `position:absolute;` +
      `left:${screenLeft}px;top:${btnY}px;width:${side}px;height:${side}px;` +
      `box-sizing:border-box;display:flex;align-items:center;justify-content:center;` +
      // Match Excel's grey button chrome; non-interactive (display only).
      `background:#f0f0f0;border:1px solid #7f7f7f;pointer-events:none;`;
    const arrow = Math.max(4, Math.round(side * 0.42));
    btn.innerHTML =
      `<svg width="${arrow}" height="${arrow}" viewBox="0 0 10 6" aria-hidden="true">` +
      `<path d="M0 0 L10 0 L5 6 Z" fill="#333"/></svg>`;
    this.host.overlayHost.appendSelection(btn);

    // Record the arrow's on-screen rect (canvasArea space) for pointer
    // hit-testing. The button element has pointer-events:none, so clicks fall
    // through to the scrollHost where the pointerdown handler tests this rect.
    this.arrowRect = { x: screenLeft, y: btnY, w: side, h: side };

    // Keep an already-open panel glued to the arrow as the grid scrolls. If the
    // active cell's validation differs from the open panel (selection moved),
    // close it instead.
    if (this.isOpen()) {
      if (this.panelKey === `${active.row}:${active.col}`) {
        this.position();
      } else {
        this.hide();
      }
    }
  }

  /** Toggle the dropdown panel for the active cell's list validation. Called
   *  from pointerdown when the arrow rect is hit. Re-clicking the same arrow
   *  closes it. */
  toggle(): void {
    const ws = this.host.worksheet();
    const active = this.host.activeCell();
    if (!ws || !active) return;
    const key = `${active.row}:${active.col}`;
    if (this.panelKey === key) {
      this.hide();
      return;
    }
    const dv = findListValidationAt(ws.dataValidations, active.row, active.col);
    if (!dv) return;
    this.hide();
    this.panelKey = key;
    void this.open(active, dv.formula1);
  }

  /** Resolve the allowed values for `formula1` (relative to the current sheet)
   *  and render them in the panel anchored below the active cell. Async because
   *  cross-sheet range references may need a lazily-parsed worksheet. */
  async open(cell: CellAddress, formula1: string | undefined): Promise<void> {
    const generation = ++this.requestGeneration;
    const workbook = this.host.workbook();
    const sheet = this.host.currentSheet();
    if (!workbook || this.host.isDestroyed()) return;
    let resolved: ResolvedList;
    try {
      resolved = await workbook.resolveValidationList(sheet, formula1);
    } catch {
      if (!this.isCurrentRequest(generation, workbook, sheet, cell)) return;
      // A resolution failure (e.g. a missing sheet) must not break the viewer;
      // fall back to disclosing the raw formula.
      resolved = { kind: 'formula', formula: formula1 ?? '' };
    }
    if (!this.isCurrentRequest(generation, workbook, sheet, cell)) return;

    this.render(resolved);
    this.position();
    this.installOutsideHandler();
  }

  private isCurrentRequest(
    generation: number,
    workbook: XlsxWorkbook,
    sheet: number,
    cell: CellAddress,
  ): boolean {
    const active = this.host.activeCell();
    return !this.host.isDestroyed()
      && generation === this.requestGeneration
      && this.host.workbook() === workbook
      && this.host.currentSheet() === sheet
      && this.panelKey === `${cell.row}:${cell.col}`
      && active?.row === cell.row
      && active?.col === cell.col;
  }

  /** Build the panel's children. Uses textContent throughout (no HTML injection
   *  from cell values). Items highlight on hover but are NOT selectable —
   *  this is a read-only viewer, so clicking a value must not change the cell. */
  private render(resolved: ResolvedList): void {
    const panel = this.panel;
    this.itemListeners.dispose();
    this.itemListeners = new ListenerScope();
    panel.textContent = '';
    if (resolved.kind === 'formula' || resolved.values.length === 0) {
      // Unresolved operand (named range / complex formula) or an empty range:
      // disclose the formula / a placeholder rather than showing a blank box.
      const note = this.host.ownerDocument.createElement('div');
      note.style.cssText = 'padding:4px 8px;color:#666;font-style:italic;white-space:pre-wrap;word-break:break-word;';
      note.textContent =
        resolved.kind === 'formula'
          ? (resolved.formula ? `= ${resolved.formula}` : '(no list)')
          : '(empty list)';
      panel.appendChild(note);
      return;
    }
    for (const value of resolved.values) {
      const item = this.host.ownerDocument.createElement('div');
      item.setAttribute('data-xlsx-validation-item', '');
      item.style.cssText = 'padding:3px 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:default;';
      item.textContent = value;
      // Hover highlight only — no click/select (read-only viewer).
      this.itemListeners.on(item, 'pointerenter', () => {
        item.style.background = '#cfe3ff';
      });
      this.itemListeners.on(item, 'pointerleave', () => {
        item.style.background = '';
      });
      panel.appendChild(item);
    }
  }

  /** Position the (already-populated, visible-or-becoming-visible) panel below
   *  the dropdown arrow / active cell using the pure geometry calculator. */
  private position(): void {
    const active = this.host.activeCell();
    if (!active) return;
    const rect = this.host.cellRect(active.row, active.col);
    if (!rect) return;
    const screenLeft = this.host.screenX(rect.x, rect.w);
    // Make it measurable off-screen first so offsetWidth/Height reflect content.
    this.panel.style.left = '-9999px';
    this.panel.style.top = '-9999px';
    this.panel.style.display = 'block';
    const pos = computeValidationPanelPosition({
      cell: { x: screenLeft, y: rect.y, w: rect.w, h: rect.h },
      panel: { w: this.panel.offsetWidth, h: this.panel.offsetHeight },
      viewport: { w: this.host.canvasArea.clientWidth, h: this.host.canvasArea.clientHeight },
      rtl: this.host.isRtl(),
    });
    this.host.overlayHost.showValidation(pos.left, pos.top);
  }

  /** Install a document-level pointerdown listener that closes the panel on a
   *  click outside it (and outside the arrow, which toggles via its own path).
   *  Removed by {@link hide}. */
  private installOutsideHandler(): void {
    if (this.outsideHandler) return;
    this.outsideHandler = (e: PointerEvent) => {
      const target = e.target as Node | null;
      if (target && this.panel.contains(target)) return; // inside panel
      // A click on the arrow is handled by the scrollHost pointerdown (toggle);
      // don't double-handle it here. Detect by hit-testing the arrow rect.
      if (this.hitsArrow(e.clientX, e.clientY)) return;
      this.hide();
    };
    // Capture phase so we see the click before it mutates selection.
    this.host.ownerDocument.addEventListener('pointerdown', this.outsideHandler, true);
  }

  /** Hide the panel and detach its outside-click listener. Called on re-click,
   *  outside click, Esc, scroll, selection change, sheet switch and destroy. */
  hide(): void {
    this.requestGeneration++;
    this.host.overlayHost.hideValidation();
    this.panelKey = null;
    if (this.outsideHandler) {
      this.host.ownerDocument.removeEventListener('pointerdown', this.outsideHandler, true);
      this.outsideHandler = null;
    }
  }

  /** Teardown: hide (dropping the document listener) and detach item hovers. */
  destroy(): void {
    this.hide();
    this.itemListeners.dispose();
  }
}
