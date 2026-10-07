import { computeFrameBox, type FrameBox } from '../frame-geometry.js';
import type { FramePr } from '../types.js';
import type { AnchorGeometryContext } from './acquisition-context.js';
import type { StoryPageFrames } from './story-page-frames.js';
import {
  wordCellOwnerContextPromotesRows,
  wordCellOwnerLeadingFrame,
  wordCellOwnerRunKey,
} from './table-compatibility.js';
import type { TableLayoutSource } from './table-source-acquisition.js';
import type {
  LayoutRect,
  SourceRef,
  TableLayout,
  TableLayoutInput,
  TableRowLayoutInput,
  TableRowOwnerCarrier,
} from './types.js';

/*
 * Cell-owner row hosts. The compatibility policies (which owning contexts
 * promote rows, which paragraph frame a row elects, the run identity key with
 * its vertical-merge and border-cap boundaries, and the host width clip) are
 * declared with their bounded evidence and normative boundary in
 * table-compatibility.ts: WORD_CELL_OWNER_ROW_CONTEXT,
 * WORD_CELL_OWNER_ROW_SELECTOR, WORD_CELL_OWNER_ROW_RUN_IDENTITY and
 * WORD_CELL_OWNER_HOST_WIDTH_CLIP. This module owns the structure around them:
 * owner contexts, carrier SourceRefs, partitioning, segment inputs, immutable
 * host geometry and the per-input segment cache. Placement, reservation,
 * splitting and clipping geometry are library choices, labeled at each rule.
 * Source rows, cells and paragraphs are never rewritten: an owner run is a
 * projection of the immutable acquired rows by reference.
 */

/**
 * The structural context that owns a table block: a direct root block of a
 * story, or a block of a table cell (a nested table, at any depth).
 */
export type TableOwnerContext =
  | Readonly<{ kind: 'story-root'; story: SourceRef['story'] }>
  | Readonly<{ kind: 'cell' }>;

export const CELL_OWNED_TABLE: TableOwnerContext = Object.freeze({ kind: 'cell' });

/**
 * Whether rows of a table in this owner context may elect a carrier at all:
 * {@link wordCellOwnerContextPromotesRows} (WORD_CELL_OWNER_ROW_CONTEXT), a
 * cell-owned table having no root story. In a non-promoting context the
 * parsed framePr stays a retained source fact and its paragraph stays
 * ordinary cell content: those tables take the ordinary table path, with no
 * owner projection and no diagnostic. Paragraph (non-table) frames are not
 * affected.
 */
export function tableRowsElectCarriers(context: TableOwnerContext): boolean {
  return wordCellOwnerContextPromotesRows(context.kind === 'story-root' ? context.story : null);
}

/**
 * The row's carrier: the frame elected by {@link wordCellOwnerLeadingFrame}
 * (WORD_CELL_OWNER_ROW_SELECTOR), sourced at block 0 of cell 0. A vMerge
 * continuation elects the frame its own first paragraph states although that
 * content is not painted; without one it stays ordinary and inherits nothing.
 * Later first-cell paragraphs keep their framePr declarations as ordinary
 * content. The caller owns the table-level gates (effective §17.4.57
 * positioning null, `word-effective-floating-table-positioning`, and
 * {@link tableRowsElectCarriers}); the paginated body and header/footer story
 * roots consume the result.
 */
export function leadingCellOwnerCarrier(
  row: TableLayoutSource['rows'][number],
  rowSource: SourceRef,
): TableRowOwnerCarrier | null {
  const framePr = wordCellOwnerLeadingFrame(row);
  if (!framePr) return null;
  return Object.freeze({
    framePr,
    source: Object.freeze({
      story: rowSource.story,
      storyInstance: rowSource.storyInstance,
      path: Object.freeze([...rowSource.path, 0, 0]),
    }),
  });
}

export interface TableOwnerSegment {
  readonly kind: 'ordinary' | 'host';
  /** First row (inclusive) and end row (exclusive) in the table input. */
  readonly rowStart: number;
  readonly rowEnd: number;
  /** Host only: the carrier shared by every electing row of the run. */
  readonly carrier?: TableRowOwnerCarrier;
}

/**
 * Partition a table into owner runs and ordinary segments, or `null` when no
 * row elects (the table keeps its single unsegmented path).
 *
 * A boundary separates exactly the adjacent rows whose run keys
 * ({@link wordCellOwnerRunKey}, WORD_CELL_OWNER_ROW_RUN_IDENTITY) differ (an
 * ordinary row has no key), so every row of a host shares its carrier.
 * Vertical merges neither move nor suppress a boundary, and no merge is
 * cleared or rewritten. ECMA-376 §17.4.84 merge integrity holds inside each
 * segment's own grid (library model): a segment is laid out as its own table,
 * so a restart closes at the segment's last row and a continue that opens a
 * segment is an empty cell with its own borders, shading and margins (its
 * row track holds them), its content not repeated
 * (`segmentOpeningLogicalRowIndex`, owned in the grid by table.ts
 * gridMergeRole).
 * Merges inside one run behave as in any table; the registered evidence
 * covers that only for a first-column merge with equal carriers.
 */
export function tableOwnerSegments(input: TableLayoutInput): readonly TableOwnerSegment[] | null {
  // Pagination asks once per fragment; the input is immutable.
  if (ownerSegmentsByInput.has(input)) return ownerSegmentsByInput.get(input)!;
  const segments = projectOwnerSegments(input);
  ownerSegmentsByInput.set(input, segments);
  return segments;
}

const ownerSegmentsByInput = new WeakMap<TableLayoutInput, readonly TableOwnerSegment[] | null>();

function projectOwnerSegments(input: TableLayoutInput): readonly TableOwnerSegment[] | null {
  const rows = input.rows;
  if (!rows.some((row) => row.ownerCarrier)) return null;
  const keys = rows.map((row) => row.ownerCarrier ? wordCellOwnerRunKey(row.ownerCarrier.framePr) : null);
  const segments: TableOwnerSegment[] = [];
  let start = 0;
  for (let index = 1; index <= rows.length; index += 1) {
    if (index < rows.length && keys[index] === keys[start]) continue;
    const carrier = rows[start]!.ownerCarrier;
    segments.push(Object.freeze({
      kind: carrier ? 'host' as const : 'ordinary' as const,
      rowStart: start,
      rowEnd: index,
      ...(carrier ? { carrier } : {}),
    }));
    start = index;
  }
  return Object.freeze(segments);
}

/** The segment owning a table-input row. Segments are contiguous and sorted,
 * so a per-fragment request costs O(log segments), not a scan. */
export function ownerSegmentAt(
  segments: readonly TableOwnerSegment[],
  rowIndex: number,
): TableOwnerSegment {
  let low = 0;
  let high = segments.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const candidate = segments[middle]!;
    if (rowIndex < candidate.rowStart) high = middle - 1;
    else if (rowIndex >= candidate.rowEnd) low = middle + 1;
    else return candidate;
  }
  throw new RangeError(`No owner segment owns table row ${rowIndex}`);
}

/**
 * One segment as its own table input. Rows are the acquired source rows by
 * reference (ids, sources, logical indexes and cell content unchanged); the
 * source table's borders are kept, so `layoutTable`'s existing outer rule
 * caps the segment with the table top/bottom (or a row exception's) and the
 * cascade does not consult insideH or a cell's conditional insideH at a run
 * edge. The border boundary behind that cap is recorded on
 * WORD_CELL_OWNER_ROW_RUN_IDENTITY; exact border geometry was not compared,
 * and the cap rule itself is this library's.
 *
 * `leadingRows` prefixes repeated source header rows for a continuation; the
 * caller decides when a fragment is a continuation (see
 * {@link ownerSegmentHeaderPrefix}). A host segment is out of ordinary flow
 * and states its carrier as `ownerHost`: the explicit contract by
 * which its owning domain (the paginated body or a header/footer story)
 * recognizes a frame owner.
 */
export function ownerSegmentInput(
  input: TableLayoutInput,
  segment: TableOwnerSegment,
  leadingRows: readonly TableRowLayoutInput[] = [],
): TableLayoutInput {
  // A continue in the first row of a later segment lost the rows above it to
  // another segment: it opens its region in this grid (see tableOwnerSegments).
  // The fact names the row by its stable logical index, which each occurrence
  // of the row keeps (a cut fragment changes only its id); the row and its
  // cells stay the source's.
  const opening = segment.rowStart > 0 ? input.rows[segment.rowStart] : undefined;
  const opensWithContinuation = opening?.cells.some((cell) => cell.verticalMerge === 'continue') ?? false;
  return Object.freeze({
    ...input,
    id: `${input.id}:owner:${segment.rowStart}-${segment.rowEnd}`,
    ordinaryFlow: segment.kind === 'host' ? false : input.ordinaryFlow,
    ...(segment.carrier ? { ownerHost: segment.carrier } : {}),
    ...(opensWithContinuation ? { segmentOpeningLogicalRowIndex: opening!.logicalRowIndex } : {}),
    rows: Object.freeze([...leadingRows, ...input.rows.slice(segment.rowStart, segment.rowEnd)]),
  });
}

/**
 * Source header rows prefixed to a segment fragment that continues the source
 * table (library policy). The source table's leading repeated-header rows that
 * precede the segment are prefixed, so with any header rows the segment itself
 * owns they form one complete leading header block of the projected input:
 * existing pagination then repeats all of them, once each, as repeated-header
 * occurrences inside the segment's owner. A fragment that resumes inside that
 * header block repeats nothing, exactly as an ordinary table resumed inside
 * its header rows does.
 */
export function ownerSegmentHeaderPrefix(
  input: TableLayoutInput,
  segment: TableOwnerSegment,
): readonly TableRowLayoutInput[] {
  let headerCount = 0;
  while (headerCount < segment.rowStart && input.rows[headerCount]?.repeatedHeader === true) {
    headerCount += 1;
  }
  return input.rows.slice(0, headerCount);
}

/**
 * An owner domain's anchor frame: `anchors` are the §17.3.1.11 bands with
 * vertical bands measured from `blockOriginPt`. Page and margin bands are
 * always the destination page's; only text bands belong to the owner (the
 * body column and cursor, or a story's column and cursor).
 */
export interface OwnerHostFrame {
  readonly anchors: AnchorGeometryContext;
  readonly blockOriginPt: number;
}

/** The host axes a page or margin frame anchor places: page coordinates,
 * which a story's band translation must not move. */
export function ownerHostPageAxes(
  framePr: Pick<FramePr, 'hAnchor' | 'vAnchor'>,
): Readonly<{ horizontal: boolean; vertical: boolean }> {
  return Object.freeze({
    horizontal: framePr.hAnchor === 'page' || framePr.hAnchor === 'margin',
    vertical: framePr.vAnchor === 'page' || framePr.vAnchor === 'margin',
  });
}

/** Whether a row carrier of this table anchors to the page or margin on
 * either axis. Only root rows elect, so nested tables never carry one. */
export function ownerCarrierAnchorsToPage(input: TableLayoutInput): boolean {
  return input.rows.some((row) => {
    if (!row.ownerCarrier) return false;
    const axes = ownerHostPageAxes(row.ownerCarrier.framePr);
    return axes.horizontal || axes.vertical;
  });
}

/**
 * Host placement P: the existing `computeFrameBox` for the run's carrier
 * with content width = the natural grid (an explicit `w` overrides only the
 * frame box) and content height = the host extent T. Page/margin anchors
 * inherit `clampAbsBoxIntoContainer` (implementation-defined, not ECMA-376);
 * a text anchor rides `cursorYPt`.
 */
export function ownerHostFrameBox(
  framePr: FramePr,
  frame: OwnerHostFrame,
  cursorYPt: number,
  gridWidthPt: number,
  extentPt: number,
): FrameBox {
  const box = computeFrameBox(
    framePr, frame.anchors, cursorYPt - frame.blockOriginPt, gridWidthPt, extentPt, 0,
  );
  const dy = frame.blockOriginPt;
  return dy === 0 ? box : { ...box, y: box.y + dy, exTop: box.exTop + dy, exBottom: box.exBottom + dy };
}

/**
 * The run carrier's frame box against a destination page's frames,
 * in the frames' coordinates: `computeFrameBox` measures its bands from the
 * page origin, which a page frame carried into a story's own coordinates
 * (a vertical section's transform) need not have at (0, 0). `contentXPt` and
 * `cursorYPt` (the text bands) are in the frames' coordinates too.
 */
export function ownerHostPageFrameBox(
  framePr: FramePr,
  frames: StoryPageFrames,
  contentXPt: number,
  contentWidthPt: number,
  cursorYPt: number,
  gridWidthPt: number,
  extentPt: number,
): FrameBox {
  const { page } = frames;
  const box = computeFrameBox(
    framePr,
    pageFrameAnchors(frames, contentXPt - page.xPt, contentWidthPt),
    cursorYPt - page.yPt,
    gridWidthPt,
    extentPt,
    0,
  );
  return page.xPt === 0 && page.yPt === 0 ? box : frameBoxAt(box, box.x + page.xPt, box.y + page.yPt);
}

/** §17.3.1.11 anchor bands of a destination page's frames, measured from
 * the page frame's origin, with the given text column (from that origin). */
function pageFrameAnchors(
  frames: StoryPageFrames,
  contentXPt: number,
  contentWidthPt: number,
): AnchorGeometryContext {
  const { page, margin } = frames;
  return {
    contentX: contentXPt,
    contentW: contentWidthPt,
    pageWidth: page.widthPt,
    pageH: page.heightPt,
    marginLeft: margin.xPt - page.xPt,
    marginRight: page.xPt + page.widthPt - margin.xPt - margin.widthPt,
    marginTop: margin.yPt - page.yPt,
    marginBottom: page.yPt + page.heightPt - margin.yPt - margin.heightPt,
  };
}

/** The same frame box moved to origin (x, y), exclusion edges included. */
export function frameBoxAt(box: FrameBox, xPt: number, yPt: number): FrameBox {
  const dx = xPt - box.x;
  const dy = yPt - box.y;
  return {
    ...box,
    x: xPt,
    y: yPt,
    exLeft: box.exLeft + dx,
    exRight: box.exRight + dx,
    exTop: box.exTop + dy,
    exBottom: box.exBottom + dy,
  };
}

/**
 * The frame box widened to the host's occupied extent (library choice): the
 * wrap exclusion spans P.y to the larger of the hRule-gated frame height and
 * the placed extent, padded by vSpace; horizontally it is the viewport (frame
 * `w`, else the grid) padded by hSpace for around/auto, as computed.
 */
export function ownerHostOccupiedBox(
  box: FrameBox,
  framePr: FramePr,
  placedExtentPt: number,
): FrameBox {
  const frameHeightPt = framePr.hRule === 'auto' ? placedExtentPt : (framePr.h ?? 0);
  const heightPt = Math.max(frameHeightPt, placedExtentPt);
  return { ...box, h: heightPt, exBottom: box.y + heightPt + framePr.vSpace };
}

function intersectRect(left: LayoutRect, right: LayoutRect): LayoutRect | null {
  const xPt = Math.max(left.xPt, right.xPt);
  const yPt = Math.max(left.yPt, right.yPt);
  const rightPt = Math.min(left.xPt + left.widthPt, right.xPt + right.widthPt);
  const bottomPt = Math.min(left.yPt + left.heightPt, right.yPt + right.heightPt);
  return rightPt > xPt && bottomPt > yPt
    ? { xPt, yPt, widthPt: rightPt - xPt, heightPt: bottomPt - yPt }
    : null;
}

/**
 * The host as painted (shared by the body and story-root domains).
 * Horizontal viewport clip of a host whose explicit frame width is narrower
 * than its natural grid (WORD_CELL_OWNER_HOST_WIDTH_CLIP in
 * table-compatibility.ts). The grid is never compressed and nothing is removed:
 * search, selection and source still own every cell. Library choices: only
 * the trailing (physical right) edge clips at `originXPt + w` — content the
 * table places at or before its origin (signed indent, half a centered
 * leading rule) is not cut; there is no vertical host clip, and the existing
 * §17.4.80 exact-row cell clips and page-end clip stay independent. RTL
 * frames were not observed.
 */
export function finishOwnerHostLayout<T extends TableLayout>(
  layout: T,
  framePr: FramePr,
  gridWidthPt: number,
  originXPt: number,
): T {
  const finished: T = Object.freeze({ ...layout, cellOwnerHost: true as const });
  if (framePr.w != null && framePr.w < gridWidthPt) {
    const ink = layout.inkBounds;
    const leadingPt = Math.min(originXPt, ink.xPt);
    const viewport: LayoutRect = {
      xPt: leadingPt,
      yPt: ink.yPt,
      widthPt: Math.max(0, originXPt + framePr.w - leadingPt),
      heightPt: ink.heightPt,
    };
    const clipBounds = layout.clipBounds
      ? intersectRect(layout.clipBounds, viewport) ?? { ...viewport, widthPt: 0 }
      : viewport;
    return Object.freeze({
      ...finished,
      clipBounds,
      inkBounds: intersectRect(ink, clipBounds) ?? { ...clipBounds },
    });
  }
  return finished;
}
