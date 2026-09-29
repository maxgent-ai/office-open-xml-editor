import type { Worksheet, XlsxChromeColors } from '../../types.js';
import { HEADER_W, HEADER_H, getGridGeometryForWorksheet } from '../../renderer.js';
import {
  buildOutlineLayout,
  toggleGroupHidden,
  levelButtonHidden,
  rowBands,
  colBands,
  summaryAfterFor,
  gutterExtentPx,
  outlineBracketSegments,
  outlineLevelButtonCenterPx,
  outlinePaneClipRect,
  OUTLINE_BUTTON_PX,
  OUTLINE_LANE_PX,
  type BandOutline,
  type OutlineGroup,
  type OutlineLayout,
  type OutlineAxis,
} from '../../outline.js';
import type { CanvasSurface } from '../sheet-surface.js';
import { ListenerScope } from './listener-scope.js';

type CellRect = { x: number; y: number; w: number; h: number };

/** Viewer state and follow-up work an outline gutter needs from its engine. */
export interface OutlineGutterHost {
  readonly gridRegion: HTMLDivElement;
  readonly canvasArea: HTMLDivElement;
  readonly surface: CanvasSurface;
  worksheet(): Worksheet | null;
  scale(): number;
  chromeColors(): XlsxChromeColors;
  /** Scrolled canvasArea rect of a cell in logical-LTR space. */
  cellRect(row: number, col: number): CellRect | null;
  /** Logical-LTR x → on-screen x (RTL mirror). */
  screenX(logicalX: number, width: number): number;
  setBandHidden(axis: OutlineAxis, index: number, hidden: boolean): void;
  setBandCollapsed(axis: OutlineAxis, index: number, collapsed: boolean): void;
  /** Rebuild geometry, overlays and the frame after a collapse-state change. */
  afterOutlineMutation(ws: Worksheet, anchor?: { axis: OutlineAxis; summary: number }): void;
}

/**
 * Row/column grouping gutters (XL4): the three gutter canvases, their layout
 * beside the grid, bracket/toggle/level-button painting and the pointer
 * hit-testing that collapses or expands groups. The collapse itself is a
 * view-only model edit delegated back to the host.
 */
export class OutlineGutter {
  /** Left gutter canvas: row group brackets + toggles (XL4). */
  readonly rowGutter: HTMLCanvasElement;
  /** Top gutter canvas: column group brackets + toggles (XL4). */
  readonly colGutter: HTMLCanvasElement;
  /** Top-left corner canvas: plain background where the two gutters meet. */
  readonly cornerGutter: HTMLCanvasElement;
  /** Cached extents (unscaled CSS px) of the current sheet's gutters; both 0 for
   *  an outline-free sheet. `w` insets canvasArea from the left, `h` from the
   *  top. */
  extent = { w: 0, h: 0 };
  /** Per-axis outline layout (group brackets + toggles) for the current sheet,
   *  recomputed on sheet switch and after each collapse/expand. `null` axis ⇒ no
   *  outlining on that axis. */
  rowOutline: OutlineLayout | null = null;
  colOutline: OutlineLayout | null = null;
  private rowOutlineBands: BandOutline[] = [];
  private colOutlineBands: BandOutline[] = [];
  private readonly listeners = new ListenerScope();

  constructor(private readonly host: OutlineGutterHost) {
    const ownerDocument = host.gridRegion.ownerDocument ?? document;
    // Absolutely positioned inside gridRegion; sized / shown per sheet in
    // `layout`. `pointer-events:auto` on the gutters so +/- toggles and level
    // buttons are clickable; they are painted on the main thread even in worker
    // mode (cheap chrome, independent of the grid bitmap).
    const gutterStyle =
      `position:absolute;top:0;left:0;z-index:3;display:none;` +
      `background:var(--ooxml-xlsx-chrome-background,#f5f5f5);`;
    this.cornerGutter = ownerDocument.createElement('canvas');
    this.cornerGutter.style.cssText = gutterStyle;
    this.cornerGutter.setAttribute('data-xlsx-outline', 'corner');
    this.colGutter = ownerDocument.createElement('canvas');
    this.colGutter.style.cssText = gutterStyle;
    this.colGutter.setAttribute('data-xlsx-outline', 'col');
    this.rowGutter = ownerDocument.createElement('canvas');
    this.rowGutter.style.cssText = gutterStyle;
    this.rowGutter.setAttribute('data-xlsx-outline', 'row');
  }

  /** Install the +/- toggle and level-bank click handling. Registered once; a
   *  no-op when a sheet has no gutter (extents 0 ⇒ detached). */
  installListeners(): void {
    this.listeners.on(this.rowGutter, 'pointerdown', (e) => this.onPointerDown(e, 'row'));
    this.listeners.on(this.colGutter, 'pointerdown', (e) => this.onPointerDown(e, 'col'));
  }

  /** Rebuild only the layout + band lists (not the stashes) after a collapse
   *  state change, so the +/- glyphs and bracket set stay in sync. Both axes are
   *  `null` (gutters collapse to 0) when the sheet has no outlining. */
  rebuild(ws: Worksheet): void {
    this.rowOutlineBands = rowBands(ws);
    this.colOutlineBands = colBands(ws);
    const rowLayout = buildOutlineLayout(this.rowOutlineBands, summaryAfterFor(ws, 'row'));
    const colLayout = buildOutlineLayout(this.colOutlineBands, summaryAfterFor(ws, 'col'));
    this.rowOutline = rowLayout.maxLevel > 0 ? rowLayout : null;
    this.colOutline = colLayout.maxLevel > 0 ? colLayout : null;
  }

  /** Size and place the three gutter canvases (corner / col / row) from the
   *  current outline, and inset canvasArea by the gutter extents. When neither
   *  axis is grouped both extents are 0 and canvasArea covers the whole region —
   *  pixel-identical to a viewer built before XL4. */
  layout(): void {
    const cs = this.host.scale();
    const gw = this.rowOutline ? Math.round(gutterExtentPx(this.rowOutline.maxLevel) * cs) : 0;
    const gh = this.colOutline ? Math.round(gutterExtentPx(this.colOutline.maxLevel) * cs) : 0;
    this.extent = { w: gw, h: gh };
    const gridRegion = this.host.gridRegion;

    // Attach the gutter canvases only while an outline exists; detach them
    // entirely for outline-free sheets. A hidden-but-attached canvas is NOT
    // neutral — DOM consumers that count/index `<canvas>` elements (e.g. the
    // layouts smoke's `page.locator('canvas').count()`) see it — so element
    // parity with the pre-outline viewer requires absence, not `display:none`.
    // The elements (and their pointer listeners) are constructed once and
    // survive detach/reattach across sheet switches.
    if (gw > 0 || gh > 0) {
      if (!this.colGutter.parentElement) {
        gridRegion.appendChild(this.colGutter);
        gridRegion.appendChild(this.rowGutter);
        gridRegion.appendChild(this.cornerGutter);
      }
    } else {
      this.colGutter.remove();
      this.rowGutter.remove();
      this.cornerGutter.remove();
    }

    // Inset canvasArea so the grid (and every geometry read that keys off its
    // client rect) starts after the gutters.
    this.host.canvasArea.style.left = `${gw}px`;
    this.host.canvasArea.style.top = `${gh}px`;

    const show = (el: HTMLCanvasElement, x: number, y: number, w: number, h: number) => {
      if (w <= 0 || h <= 0) { el.style.display = 'none'; return; }
      el.style.display = 'block';
      el.style.left = `${x}px`;
      el.style.top = `${y}px`;
      el.style.width = `${w}px`;
      el.style.height = `${h}px`;
    };
    const regionW = gridRegion.clientWidth;
    const regionH = gridRegion.clientHeight;
    // Corner holds the numbered level buttons; only meaningful where both a
    // horizontal and vertical gutter exist, but we always paint it to cover the
    // intersection so the two strips meet cleanly.
    show(this.cornerGutter, 0, 0, gw, gh);
    show(this.colGutter, gw, 0, Math.max(0, regionW - gw), gh);
    show(this.rowGutter, 0, gh, gw, Math.max(0, regionH - gh));
  }

  /** Paint all visible gutter strips for the current scroll offset. Called at the
   *  end of every grid render so the brackets track scroll / zoom exactly. */
  render(): void {
    const ws = this.host.worksheet();
    if (!ws) return;
    if (this.extent.h > 0 && this.colOutline) this.paintAxisGutter('col');
    if (this.extent.w > 0 && this.rowOutline) this.paintAxisGutter('row');
    if (this.extent.w > 0 || this.extent.h > 0) this.paintCornerGutter();
  }

  /** Draw one axis's group brackets and +/- toggles into its gutter canvas,
   *  aligned to the on-screen band positions via the host's cell rects. */
  private paintAxisGutter(axis: OutlineAxis): void {
    const ws = this.host.worksheet();
    if (!ws) return;
    const cs = this.host.scale();
    const colors = this.host.chromeColors();
    const isRow = axis === 'row';
    const canvas = isRow ? this.rowGutter : this.colGutter;
    const layout = isRow ? this.rowOutline : this.colOutline;
    if (!layout) return;
    const cssW = parseFloat(canvas.style.width) || 0;
    const cssH = parseFloat(canvas.style.height) || 0;
    if (cssW <= 0 || cssH <= 0) return;
    // Backing-store size at DPR; CSS size stays as laid out.
    const dpr = this.host.surface.sizeCanvas(canvas, cssW, cssH);
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    ctx.fillStyle = colors.background ?? '#f5f5f5';
    ctx.fillRect(0, 0, cssW, cssH);

    const lanePx = OUTLINE_LANE_PX * cs;
    // The gutter canvas's cross-axis origin (0) sits at the grid's cell-area
    // origin: for the row gutter, y=0 aligns with the top of the row header +
    // gutter; cellRect returns coordinates in canvasArea space, which is offset
    // from the gutter canvas by exactly `extent.h` (col gutter is above). The
    // gutter canvas top is at gridRegion y = extent.h, and canvasArea top is
    // also at extent.h — so a band's canvasArea-space y maps 1:1 to gutter-canvas
    // y. Likewise x for the col gutter (offset by extent.w).
    ctx.strokeStyle = colors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillStyle = colors.text ?? '#404040';

    // Outline gutter geometry participates in the same header/frozen-pane split
    // as the worksheet canvas. Clip every logical run to its own pane so a
    // scrolled detail rail cannot leak upward into the column-letter header or
    // through the frozen-row boundary (and mirror the equivalent rule for RTL
    // frozen columns).
    const geometry = getGridGeometryForWorksheet(ws);
    const effective = geometry.effectiveFrozenBands({
      scale: cs,
      width: this.host.canvasArea.clientWidth,
      height: this.host.canvasArea.clientHeight,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      rows: ws.freezeRows ?? 0,
      cols: ws.freezeCols ?? 0,
    });
    const axes = geometry.axesAtScale(cs);
    const frozenBandCount = isRow ? effective.rows : effective.cols;
    const frozenExtent = isRow
      ? axes.row.offsetOf(effective.rows + 1)
      : axes.col.offsetOf(effective.cols + 1);
    const headerExtent = (isRow ? HEADER_H : HEADER_W) * cs;
    const paneClip = (start: number, end: number) => outlinePaneClipRect(
      axis,
      start,
      end,
      frozenBandCount,
      headerExtent,
      frozenExtent,
      cssW,
      cssH,
      !isRow && ws.rightToLeft === true,
    );
    const clipContext = (start: number, end: number): boolean => {
      const clip = paneClip(start, end);
      if (clip.w <= 0 || clip.h <= 0) return false;
      ctx.save();
      ctx.beginPath();
      ctx.rect(clip.x, clip.y, clip.w, clip.h);
      ctx.clip();
      return true;
    };

    for (const g of layout.groups) {
      // Lane index for this level: lane 0 is the outermost (level 1). Buttons and
      // the outermost bracket sit nearest the grid edge? Excel draws level 1 in
      // the lane FARTHEST from the grid, deeper levels closer. We place level L in
      // lane (L-1) counted from the sheet-far edge.
      const laneFromFar = g.level - 1;
      const laneCenterCross = (laneFromFar + 0.5) * lanePx;

      // Detail run extent along the band axis, from on-screen cell rects.
      const startRect = isRow ? this.host.cellRect(g.start, 1) : this.host.cellRect(1, g.start);
      const endRect = isRow ? this.host.cellRect(g.end, 1) : this.host.cellRect(1, g.end);
      if (!startRect || !endRect) continue;
      const a = isRow ? startRect.y : this.host.screenX(startRect.x, startRect.w);
      const b = isRow ? endRect.y + endRect.h : this.host.screenX(endRect.x, endRect.w) + endRect.w;
      const runStart = Math.min(a, b);
      const runEnd = Math.max(a, b);

      // A collapsed group's detail run is hidden (zero visible extent) — Excel
      // draws only the +/- toggle, no bracket. Skip the bracket when the run has
      // negligible length.
      if (!g.collapsed && runEnd - runStart > 1) {
        if (clipContext(g.start, g.end)) {
          ctx.beginPath();
          for (const segment of outlineBracketSegments(axis, laneCenterCross, a, b, lanePx)) {
            ctx.moveTo(segment.x1, segment.y1);
            ctx.lineTo(segment.x2, segment.y2);
          }
          ctx.stroke();
          ctx.restore();
        }
      }

      // +/- toggle box on the summary band.
      if (g.summary != null) {
        const sRect = isRow ? this.host.cellRect(g.summary, 1) : this.host.cellRect(1, g.summary);
        if (sRect) {
          const along = isRow
            ? sRect.y + sRect.h / 2
            : this.host.screenX(sRect.x, sRect.w) + sRect.w / 2;
          if (clipContext(g.summary, g.summary)) {
            this.drawToggleBox(ctx, isRow ? laneCenterCross : along, isRow ? along : laneCenterCross, g.collapsed, cs);
            ctx.restore();
          }
        }
      }
    }

    // Numbered level buttons (1..maxLevel+1), one per lane, in this gutter's
    // header strip: the row bank sits beside the column-letter header (the
    // gutter's top HEADER_H band — no bracket ever draws there because band
    // y-coordinates start at the header edge), the column bank above the
    // row-number header (leftmost HEADER_W band). Placing each bank in its own
    // gutter (Excel's layout) keeps the two banks from ever sharing a cell —
    // the old corner placement collided at the shared bottom-right lane and
    // made the row expand-all button unreachable.
    const bankCross = isRow ? (HEADER_H * cs) / 2 : (HEADER_W * cs) / 2;
    for (let l = 1; l <= layout.maxLevel + 1; l++) {
      const buttonCenter = outlineLevelButtonCenterPx(l) * cs;
      if (buttonCenter + (OUTLINE_BUTTON_PX * cs) / 2 > (isRow ? cssW : cssH) + 0.5) break;
      this.drawLevelButton(
        ctx,
        isRow ? buttonCenter : bankCross,
        isRow ? bankCross : buttonCenter,
        String(l),
        cs,
      );
    }

    // Paint the pane separator last so it visibly cuts the outline rail at the
    // same coordinate as the main grid's separator. This also extends the line
    // through the outline gutter, making the frozen-row boundary continuous
    // from the gutter through the row-number header and cells.
    if (frozenBandCount > 0) {
      const divider = isRow
        ? headerExtent + frozenExtent
        : ws.rightToLeft === true
          ? cssW - headerExtent - frozenExtent
          : headerExtent + frozenExtent;
      ctx.save();
      ctx.strokeStyle = colors.border ?? '#7a7a7a';
      ctx.lineWidth = 0.5;
      ctx.beginPath();
      if (isRow) {
        ctx.moveTo(0, divider);
        ctx.lineTo(cssW, divider);
      } else {
        ctx.moveTo(divider, 0);
        ctx.lineTo(divider, cssH);
      }
      ctx.stroke();
      ctx.restore();
    }
  }

  /** Draw a small square +/- toggle centered at (cx, cy) in gutter-canvas CSS px. */
  private drawToggleBox(
    ctx: CanvasRenderingContext2D,
    cx: number,
    cy: number,
    collapsed: boolean,
    cs: number,
  ): void {
    const colors = this.host.chromeColors();
    const s = Math.round(9 * cs);
    const x = Math.round(cx - s / 2);
    const y = Math.round(cy - s / 2);
    ctx.save();
    ctx.fillStyle = colors.surface ?? '#ffffff';
    ctx.strokeStyle = colors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeStyle = colors.text ?? '#404040';
    ctx.beginPath();
    // horizontal stroke (present for both + and -)
    ctx.moveTo(x + 2.5, y + s / 2 + 0.5);
    ctx.lineTo(x + s - 1.5, y + s / 2 + 0.5);
    if (collapsed) {
      // vertical stroke makes it a "+"
      ctx.moveTo(x + s / 2 + 0.5, y + 2.5);
      ctx.lineTo(x + s / 2 + 0.5, y + s - 1.5);
    }
    ctx.stroke();
    ctx.restore();
  }

  /** Draw one numbered level button centered at (cx, cy) in gutter-canvas CSS
   *  px. Shared by the row bank (in the row gutter's top strip) and the column
   *  bank (in the column gutter's left strip). */
  private drawLevelButton(
    ctx: CanvasRenderingContext2D,
    cx: number,
    cy: number,
    label: string,
    cs: number,
  ): void {
    const colors = this.host.chromeColors();
    const s = Math.round(OUTLINE_BUTTON_PX * cs);
    const x = Math.round(cx - s / 2);
    const y = Math.round(cy - s / 2);
    ctx.save();
    ctx.font = `${Math.round(9 * cs)}px sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillStyle = colors.surface ?? '#ffffff';
    ctx.strokeStyle = colors.border ?? '#808080';
    ctx.lineWidth = 1;
    ctx.fillRect(x + 0.5, y + 0.5, s, s);
    ctx.strokeRect(x + 0.5, y + 0.5, s, s);
    ctx.fillStyle = colors.text ?? '#404040';
    ctx.fillText(label, cx, cy + 0.5);
    ctx.restore();
  }

  /** Paint the corner (intersection of the two gutters) as plain background.
   *  The numbered level banks live in each axis gutter's own header strip
   *  (see paintAxisGutter), so the corner carries no interactive content. */
  private paintCornerGutter(): void {
    const canvas = this.cornerGutter;
    const cssW = parseFloat(canvas.style.width) || 0;
    const cssH = parseFloat(canvas.style.height) || 0;
    if (cssW <= 0 || cssH <= 0) { return; }
    const dpr = this.host.surface.sizeCanvas(canvas, cssW, cssH);
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, cssW, cssH);
    ctx.fillStyle = this.host.chromeColors().background ?? '#f5f5f5';
    ctx.fillRect(0, 0, cssW, cssH);
  }

  /** Handle a click in a row/col gutter: hit-test the +/- toggles and toggle the
   *  matching group's collapse state. */
  private onPointerDown(e: PointerEvent, axis: OutlineAxis): void {
    const ws = this.host.worksheet();
    if (!ws) return;
    const isRow = axis === 'row';
    const layout = isRow ? this.rowOutline : this.colOutline;
    if (!layout) return;
    const canvas = isRow ? this.rowGutter : this.colGutter;
    const rect = canvas.getBoundingClientRect();
    const px = e.clientX - rect.left;
    const py = e.clientY - rect.top;
    const cs = this.host.scale();
    const lanePx = OUTLINE_LANE_PX * cs;
    const hitR = 7 * cs; // generous grab radius around a +/- button center

    // Numbered level bank first: it lives in this gutter's header strip (row
    // bank beside the column-letter header, column bank above the row-number
    // header — mirrors paintAxisGutter), where no +/- toggle can be.
    const bankCross = isRow ? (HEADER_H * cs) / 2 : (HEADER_W * cs) / 2;
    const inBankStrip = (isRow ? py : px) <= (isRow ? HEADER_H : HEADER_W) * cs;
    if (inBankStrip) {
      for (let l = 1; l <= layout.maxLevel + 1; l++) {
        const buttonCenter = outlineLevelButtonCenterPx(l) * cs;
        const cx = isRow ? buttonCenter : bankCross;
        const cy = isRow ? bankCross : buttonCenter;
        const buttonHitR = (OUTLINE_BUTTON_PX * cs) / 2;
        if (Math.abs(px - cx) <= buttonHitR && Math.abs(py - cy) <= buttonHitR) {
          e.preventDefault();
          this.applyLevelButton(l, axis);
          return;
        }
      }
      return; // header strip carries no toggles — don't fall through
    }

    for (const g of layout.groups) {
      if (g.summary == null) continue;
      const laneCenterCross = (g.level - 1 + 0.5) * lanePx;
      const sRect = isRow ? this.host.cellRect(g.summary, 1) : this.host.cellRect(1, g.summary);
      if (!sRect) continue;
      const along = isRow
        ? sRect.y + sRect.h / 2
        : this.host.screenX(sRect.x, sRect.w) + sRect.w / 2;
      const cx = isRow ? laneCenterCross : along;
      const cy = isRow ? along : laneCenterCross;
      if (Math.abs(px - cx) <= hitR && Math.abs(py - cy) <= hitR) {
        e.preventDefault();
        this.applyGroupToggle(g, axis);
        return;
      }
    }
  }

  /** Flip a single group's collapse state in the in-memory model, then rebuild
   *  the outline + repaint. View-only: the file is never written. */
  applyGroupToggle(group: OutlineGroup, axis: OutlineAxis): void {
    const ws = this.host.worksheet();
    if (!ws) return;
    const bands = axis === 'row' ? this.rowOutlineBands : this.colOutlineBands;
    const { hide, show, nowCollapsed } = toggleGroupHidden(group, bands);
    for (const i of hide) this.host.setBandHidden(axis, i, true);
    for (const i of show) this.host.setBandHidden(axis, i, false);
    // Reflect the new collapsed state on the summary band so the next toggle
    // reads the correct direction and the +/- glyph flips.
    if (group.summary != null) this.host.setBandCollapsed(axis, group.summary, nowCollapsed);
    // Collapsing removes the detail bands before the summary. Anchor that
    // surviving summary band at the viewport start after geometry has been
    // rebuilt; otherwise the browser clamps the shortened scroll extent and
    // leaves an unrelated partial row at the top.
    this.host.afterOutlineMutation(
      ws,
      nowCollapsed && group.summary != null ? { axis, summary: group.summary } : undefined,
    );
  }

  /** Collapse/expand the whole sheet to `level` on one axis. */
  applyLevelButton(level: number, axis: OutlineAxis): void {
    const ws = this.host.worksheet();
    if (!ws) return;
    const bands = axis === 'row' ? this.rowOutlineBands : this.colOutlineBands;
    const { hide, show } = levelButtonHidden(bands, level);
    for (const i of hide) this.host.setBandHidden(axis, i, true);
    for (const i of show) this.host.setBandHidden(axis, i, false);
    // Update each group's summary-band collapsed flag from the new state: a group
    // at lane L is collapsed exactly when its detail (level >= L) is now hidden,
    // i.e. `L >= level`. Driving this off the layout's groups (rather than the
    // band list) also reaches level-0 summary bands, which are not in `bands`.
    const layout = axis === 'row' ? this.rowOutline : this.colOutline;
    if (layout) {
      for (const g of layout.groups) {
        if (g.summary != null) this.host.setBandCollapsed(axis, g.summary, g.level >= level);
      }
    }
    this.host.afterOutlineMutation(ws);
  }

  /** Teardown: detach listeners and forget the current sheet's layout. */
  destroy(): void {
    this.listeners.dispose();
    this.rowOutlineBands = [];
    this.colOutlineBands = [];
    this.rowOutline = null;
    this.colOutline = null;
  }
}
