// Text-frame / drop-cap placement geometry (ECMA-376 §17.3.1.11 `<w:framePr>`).
//
// Pure placement math: given a `<w:framePr>` and the section geometry on
// `AnchorGeometryContext`, resolve the frame box (points) and the wrap-exclusion
// FloatRect it pushes onto `state.floats`. Extracted from renderer.ts so the
// resolve logic can be unit-reasoned in isolation (see frame-geometry.test.ts /
// measure-column-geometry.test.ts).
//
// This module is the shared base for floating-table placement
// (float-table-geometry.ts), which reuses frameXContainer / frameYContainer /
// resolveAlignedPosH / resolveAlignedPosV (the anchor/alignment semantics line
// up 1:1 between a text frame and a floating table) and pushFloatRect (the
// single source of exclusion-rect construction).

import type { FramePr } from './types.js';
import type {
  AnchorGeometryContext,
  FloatRegistrationState,
} from './layout/acquisition-context.js';
import type { FloatRect } from './float-layout.js';
import {
  FLOAT_OVERLAP_EPS,
  FLOAT_PAGE_RIGHT_SLACK,
  drawingMLAvoidance,
  frameAvoidance,
  floatRectParticipant,
  floatingTableAvoidance,
  resolveFloatPlacement,
  type FloatPlacementParticipant,
} from './layout/floats.js';

/** Resolved point-space geometry of a `<w:framePr>` text frame. Exported for
 *  unit tests only (the table-driven frame-geometry assertions) — not part of
 *  the package API. */
export interface FrameBox {
  /** Drawing origin of the frame content (text area top-left). */
  x: number;
  y: number;
  /** Frame content width / height. */
  w: number;
  h: number;
  /** Padded exclusion rect for the wrap FloatRect (frame + hSpace/vSpace). */
  exLeft: number;
  exRight: number;
  exTop: number;
  exBottom: number;
  /** False for non-owning paragraphs in one grouped frame. */
  registerExclusion?: boolean;
  /** Stable retained identity for the group's one wrap exclusion. */
  exclusionId?: string;
}

/**
 * Horizontal container band for a frame's hAnchor (ECMA-376 §17.3.1.11 /
 * §17.18.35). This is a SEPARATE relativeFrom set from DrawingML's
 * §20.4.3 (so `xContainer` in anchor-geometry is intentionally not reused):
 *   - "text"   → the COLUMN text margin the anchor paragraph sits in
 *                (state.contentX..contentX+contentW). This keeps a drop cap
 *                inside its own newspaper column (#513 per-section columns).
 *   - "margin" → the page content margin (marginLeft..pageWidth-marginRight).
 *   - "page"   → the physical page edges (0..pageWidth).
 * All values are in points.
 */
export function frameXContainer(
  hAnchor: string,
  state: AnchorGeometryContext,
): { left: number; right: number } {
  switch (hAnchor) {
    case 'margin':
      return { left: state.marginLeft, right: state.pageWidth - state.marginRight };
    case 'page':
      return { left: 0, right: state.pageWidth };
    case 'text':
    case 'column':
    default:
      // "text" anchors against the current COLUMN band so a frame in a multi-
      // column section stays inside its column.
      return { left: state.contentX, right: state.contentX + state.contentW };
  }
}

/**
 * Vertical container band for a frame's vAnchor (ECMA-376 §17.3.1.11 /
 * §17.18.100). Symmetric with {@link frameXContainer}: ST_YAlign positions the
 * frame relative to the ANCHOR OBJECT (this band), not the physical page
 * (§22.9.2.20: "this relative position is specified relative to the vertical
 * anchor"). `paraTop` is the anchor paragraph's point-space text-area top and
 * `contentH` its point-space frame-content height, used only for the "text"
 * band end.
 *   - "page"   → [0, pageH]: the physical page edges (§17.18.100 page = "the
 *                location of the edge of the page").
 *   - "margin" → [marginTop, pageH−marginBottom]: the text margins (§17.18.100
 *                margin = "the location of the text margin").
 *   - "text"   → [paraTop, paraTop+contentH]: the anchor paragraph's text
 *                extents (§17.18.100 text = "the top edge of the text in the
 *                anchor paragraph"). Relative positioning (yAlign) is not
 *                allowed for "text" (§17.3.1.11 yAlign), so only `start` is ever
 *                consumed (as the base for the absolute y offset).
 * All values are in points.
 */
export function frameYContainer(
  vAnchor: string,
  paraTop: number,
  contentH: number,
  state: AnchorGeometryContext,
): { start: number; end: number } {
  switch (vAnchor) {
    case 'margin':
      return { start: state.marginTop, end: state.pageH - state.marginBottom };
    case 'page':
      return { start: 0, end: state.pageH };
    case 'text':
    default:
      return { start: paraTop, end: paraTop + contentH };
  }
}

/**
 * Resolve a horizontal aligned position in points for a frame (xAlign,
 * §17.3.1.11) or a floating table (tblpXSpec, §17.4.57). Both use the same
 * ST_XAlign vocabulary against a container band [containerLeft, containerRight]:
 *   center          → box centred in the band
 *   right / outside  → box flush to the band's right edge
 *   left / inside / * → box flush to the band's left edge (the default)
 * Shared by {@link computeFrameBox} and computeFloatTableBox so the two
 * stay byte-identical.
 */
export function resolveAlignedPosH(
  spec: string,
  containerLeft: number,
  containerRight: number,
  size: number,
): number {
  switch (spec) {
    case 'center':
      return containerLeft + (containerRight - containerLeft - size) / 2;
    case 'right':
    case 'outside':
      return containerRight - size;
    case 'left':
    case 'inside':
    default:
      return containerLeft;
  }
}

/**
 * Resolve a vertical aligned position in points for a frame (yAlign,
 * §17.3.1.11) or a floating table (tblpYSpec, §17.4.57). Both use the same
 * ST_YAlign vocabulary, measured against the vAnchor BAND `[band.start,
 * band.end]` (the anchor object, §22.9.2.20) — symmetric with
 * {@link resolveAlignedPosH}:
 *   center           → box centred within the band
 *   bottom / outside  → box flush to the band's end (bottom) edge
 *   top / inside / inline / * → box flush to the band's start (top) edge (default)
 * Callers gate this on vAnchor!=='text' (relative vertical positioning is not
 * allowed there, §17.3.1.11 yAlign). Shared by {@link computeFrameBox} and
 * computeFloatTableBox.
 */
export function resolveAlignedPosV(
  spec: string,
  band: { start: number; end: number },
  size: number,
): number {
  switch (spec) {
    case 'center':
      return band.start + (band.end - band.start - size) / 2;
    case 'bottom':
    case 'outside':
      return band.end - size;
    case 'top':
    case 'inside':
    case 'inline':
    default:
      return band.start;
  }
}

/**
 * Implementation-defined overflow policy for an absolutely positioned
 * vAnchor=page/margin box: shift an overflowing bottom edge back into its
 * vertical container. ECMA-376 does not define this clamp. Unlike a text-anchored
 * box, an absolute page/margin position cannot be repaired by moving the flow
 * cursor, so this local geometry policy keeps the full box reachable.
 *
 *   y = max(containerStart, containerEnd − boxH)
 *
 * The floor is `containerStart` (container top): a box TALLER than its container
 * pins to the top and is allowed to overflow the bottom (clamping to
 * `end − boxH < start` would push it ABOVE the container top, which is worse). For
 * vAnchor="page" the container is the physical page [0, pageH]. The
 * vAnchor="margin" target is explicitly implementation-defined because
 * ECMA-376 does not specify the overflow clamp: use the owning margin band's
 * end, symmetric with the page case. Callers pass the same
 * `frameYContainer(vAnchor,…)` band used for placement, so the target remains
 * consistent with the anchor semantics.
 *
 * Only meaningful for vAnchor=page/margin: the caller gates on that (vAnchor=text
 * is handled by pagination, and its band start/end ride the flow cursor, so this
 * clamp must not run there). Idempotent for a box that already fits (y unchanged).
 */
export function clampAbsBoxIntoContainer(
  y: number,
  boxH: number,
  band: { start: number; end: number },
): number {
  if (y + boxH <= band.end) return y; // already inside — no-op (common case).
  return Math.max(band.start, band.end - boxH);
}

/**
 * Resolve a frame's box in points. `paraTop` is the in-flow top of the frame
 * paragraph (post-spaceBefore). `contentW`/`contentH` are the frame content's
 * measured natural size (points); `anchorLineHPt` is one line height of the
 * following non-frame (anchor) paragraph, used to size a drop cap by `lines`.
 *
 * Exported for unit tests only (frame-geometry table) — not package API.
 */
export function computeFrameBox(
  fp: FramePr,
  state: AnchorGeometryContext,
  paraTop: number,
  contentW: number,
  contentH: number,
  anchorLineHPt: number,
): FrameBox {
  const isDropCap = fp.dropCap === 'drop' || fp.dropCap === 'margin';

  const hx = frameXContainer(fp.hAnchor, state);
  // Vertical band of the vAnchor (the "anchor object", §22.9.2.20). The "text"
  // band end uses the frame content height; yAlign is gated out for "text" so
  // only band.start (= paraTop) is consumed there.
  const vBand = frameYContainer(fp.vAnchor, paraTop, contentH, state);

  // Frame width: explicit `w` (exact) else natural content width (§17.3.1.11 w).
  const frameW = fp.w != null ? fp.w : contentW;

  // Frame height. For a drop cap the height is `lines` × the anchor paragraph's
  // line height (§17.3.1.11 lines: "the height of the drop cap is the first N
  // lines of the anchor paragraph"; y/yAlign are ignored). For a generic frame
  // hRule gates h: exact = h, atLeast = max(h, content), auto = content.
  let frameH: number;
  if (isDropCap) {
    frameH = Math.max(1, fp.lines) * anchorLineHPt;
  } else {
    const hPt = fp.h ?? 0;
    frameH =
      fp.hRule === 'exact'
        ? hPt
        : fp.hRule === 'atLeast'
          ? Math.max(hPt, contentH)
          : contentH;
  }

  // Horizontal placement.
  //   dropCap="drop"   → inside the column/text margin (frame at band left).
  //   dropCap="margin" → outside the margin (frame left = band left − frameW).
  //   generic frame    → xAlign (left/center/right/inside/outside) supersedes x;
  //                      else absolute x offset from the hAnchor's left edge.
  let frameX: number;
  if (fp.dropCap === 'drop') {
    frameX = hx.left;
  } else if (fp.dropCap === 'margin') {
    frameX = hx.left - frameW;
  } else if (fp.xAlign) {
    frameX = resolveAlignedPosH(fp.xAlign, hx.left, hx.right, frameW);
  } else {
    // §17.3.1.11 x: absolute signed offset from the hAnchor left edge.
    frameX = hx.left + (fp.x ?? 0);
  }

  // Vertical placement. For a drop cap, y/yAlign are ignored: the cap sits at
  // the anchor paragraph top (§17.3.1.11 lines). Otherwise yAlign supersedes y
  // (ignored when vAnchor="text" — relative positioning is not allowed there,
  // §17.3.1.11 yAlign), else absolute y offset from the vAnchor edge.
  let frameY: number;
  if (isDropCap) {
    frameY = vBand.start;
  } else if (fp.yAlign && fp.vAnchor !== 'text') {
    frameY = resolveAlignedPosV(fp.yAlign, vBand, frameH);
  } else {
    // §17.3.1.11 y: absolute signed offset from the vAnchor band start.
    frameY = vBand.start + (fp.y ?? 0);
  }

  // Apply the implementation-defined absolute-box overflow policy. A
  // vAnchor="text" frame is excluded: its band moves with flow and pagination
  // owns keep-with-anchor behavior. See clampAbsBoxIntoContainer.
  if (fp.vAnchor === 'page' || fp.vAnchor === 'margin') {
    frameY = clampAbsBoxIntoContainer(frameY, frameH, vBand);
  }

  // Exclusion padding: hSpace L/R applies only with wrap="around" (§17.3.1.11
  // hSpace); vSpace top/bottom always.
  const hSpacePt = fp.wrap === 'around' || fp.wrap === 'auto' ? fp.hSpace : 0;
  const vSpacePt = fp.vSpace;

  return {
    x: frameX,
    y: frameY,
    w: frameW,
    h: frameH,
    exLeft: frameX - hSpacePt,
    exRight: frameX + frameW + hSpacePt,
    exTop: frameY - vSpacePt,
    exBottom: frameY + frameH + vSpacePt,
  };
}

/** Options for {@link pushFloatRect}: the resolved point-space image/float box
 *  (x,y,w,h), its dist* padding (dl,dr,dt,db), and the FloatRect descriptors. */
export interface PushFloatOpts {
  x: number;
  y: number;
  w: number;
  h: number;
  dl: number;
  dr: number;
  dt: number;
  db: number;
  mode: 'square' | 'topAndBottom';
  side: string;
  imageKey: string;
  paraId: number;
  kind: FloatRect['kind'];
  /** Required at runtime for table entries; the resulting FloatRect and typed
   * placement participant make the fact structurally mandatory. */
  tableOverlap?: 'never' | 'overlap';
  allowOverlap?: boolean;
  avoidOverlap: boolean;
}

/**
 * Build a wrap-exclusion {@link FloatRect} from a resolved box + dist padding,
 * optionally running overlap avoidance first, push it onto `state.floats`, and
 * return it. Single source of the `xLeft = x − dl, xRight = x + w + dr,
 * yTop = y − dt, yBottom = y + h + db` exclusion-rect construction shared by
 * registerFrameFloat / registerTableFloat / registerImageFloat /
 * registerShapeFloat (the `dist*` fields carry dl/dr/dt/db verbatim for
 * re-seating). The returned ref exposes the exact registered transport record
 * to acquisition callers that need its overlap-resolved position.
 */
export function pushFloatRect(state: FloatRegistrationState, o: PushFloatOpts): FloatRect {
  if (o.kind === 'table' && o.tableOverlap === undefined) {
    throw new Error('Floating-table transport omitted tblOverlap');
  }
  let px = o.x;
  let py = o.y;
  if (o.avoidOverlap) {
    const core = {
      occurrenceId: 'display-moving-float',
      paragraphId: o.paraId,
      bounds: { xPt: px, yPt: py, widthPt: o.w, heightPt: o.h },
      exclusionBounds: {
        xPt: px - o.dl,
        yPt: py - o.dt,
        widthPt: o.w + o.dl + o.dr,
        heightPt: o.h + o.dt + o.db,
      },
    };
    const moving: FloatPlacementParticipant = o.kind === 'table'
      ? { ...core, kind: 'table', tableOverlap: o.tableOverlap! }
      : { ...core, kind: o.kind === 'frame' ? 'frame' : 'drawingml' };
    const resolved = resolveFloatPlacement({
      moving,
      blockers: state.floats.map(floatRectParticipant),
      avoidance: o.kind === 'table'
        ? floatingTableAvoidance(o.tableOverlap!, o.paraId)
        : o.kind === 'frame'
          ? frameAvoidance(o.paraId)
          : drawingMLAvoidance(o.allowOverlap ?? true),
      rightBoundaryPt: state.pageWidth,
      overlapEpsilonPt: FLOAT_OVERLAP_EPS,
      rightBoundarySlackPt: FLOAT_PAGE_RIGHT_SLACK,
    });
    px = resolved.bounds.xPt;
    py = resolved.bounds.yPt;
  }
  const core = {
    mode: o.mode,
    imageKey: o.imageKey,
    imageX: px,
    imageY: py,
    imageW: o.w,
    imageH: o.h,
    xLeft: px - o.dl,
    xRight: px + o.w + o.dr,
    yTop: py - o.dt,
    yBottom: py + o.h + o.db,
    side: o.side,
    distLeft: o.dl,
    distRight: o.dr,
    distTop: o.dt,
    distBottom: o.db,
    paraId: o.paraId,
  };
  const rect: FloatRect = o.kind === 'table'
    ? { ...core, kind: 'table', tableOverlap: o.tableOverlap! }
    : { ...core, kind: o.kind };
  state.floats.push(rect);
  return rect;
}

/**
 * The exclusion a framePr wrap imposes (mapping documented on
 * {@link registerFrameFloat}). This is §17.18.104 ST_Wrap of w:framePr, not
 * DrawingML wrapNone (§20.4.2.15), whose text does overlap its object.
 */
export function frameWrapExclusionMode(
  fp: Pick<FramePr, 'wrap'>,
): 'square' | 'topAndBottom' {
  return fp.wrap === 'none' || fp.wrap === 'notBeside' ? 'topAndBottom' : 'square';
}

/**
 * Push the wrap-exclusion FloatRect for a resolved frame box onto
 * `state.floats` so following body text flows around the frame. No-op for
 * a degenerate (zero-area) box. Shared by the renderer (after drawing) and
 * the paginator (so the anchor paragraph's measured height accounts for the
 * wrap). The exclusion x-range is COLUMN-relative (built in frameXContainer
 * from state.contentX/contentW for hAnchor="text"), so
 * resolveLineFloatWindow only constrains the matching column (#513).
 *
 * Wrap-mode → FloatRect mapping (ECMA-376 §17.18.104):
 *   none, notBeside → topAndBottom: neither allows text to wrap around the
 *               frame; text resumes on the next line that does not intersect
 *               the frame's extents.
 *   around / auto → square side wrap. `word-frame-auto-wrap-around` records
 *               the application-defined mapping of auto to around.
 *   tight / through → a frame is a rectangle, so contour wrapping collapses to
 *               a square wrap (no contour follow for a rectangular frame).
 *
 * Exported for unit tests only (frame-geometry table) — not package API.
 */
export function registerFrameFloat(box: FrameBox, fp: FramePr, state: FloatRegistrationState): void {
  if (box.registerExclusion === false) return;
  const mode = frameWrapExclusionMode(fp);
  if (box.w <= 0 || box.h <= 0) return;

  const paraId = state.floatParaSeq++;
  // dist padding recovered from the box's pre-computed exclusion edges so the
  // unified builder reproduces xLeft=box.exLeft etc. exactly.
  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    dl: box.x - box.exLeft,
    dr: box.exRight - (box.x + box.w),
    dt: box.y - box.exTop,
    db: box.exBottom - (box.y + box.h),
    kind: 'frame',
    mode,
    // A drop cap sits at the column's left edge, so text wraps only to its
    // RIGHT. A generic frame may sit anywhere, so text wraps on both sides
    // (resolveLineFloatWindow then takes the widest free gap around it).
    side: fp.dropCap === 'drop' || fp.dropCap === 'margin' ? 'right' : 'bothSides',
    imageKey: box.exclusionId ?? '',
    paraId,
    avoidOverlap: false, // frames opt out of overlap re-seating.
  });
}
