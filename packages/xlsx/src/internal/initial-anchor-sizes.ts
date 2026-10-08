/**
 * Prepared-initial display sizes for XML `twoCellAnchor editAs="oneCell"`
 * pictures and shape groups (#1713). Internal integration seam, not public
 * API: nothing here is exported from the package entry points.
 *
 * Normative facts (ECMA-376 Part 1): a `twoCellAnchor`'s from/to markers are
 * its initial display rectangle (§20.5.2.33), and `editAs="oneCell"` means
 * later band edits move the object without resizing it (§20.5.3.2). A
 * `oneCellAnchor` carries its own extent (§20.5.2.24) and is not captured.
 *
 * Library policy:
 * - "Initial" means the viewer-prepared display: host fonts bound (pinned
 *   maximum digit width) and automatic row heights prepared, BEFORE any
 *   manual band edit. Capture reads that state through the worksheet's
 *   existing GridGeometry at the viewer's first prepared scale. Grids round
 *   each band's scaled pixels, so that scale influences the reference; the
 *   captured pixels are divided by it and stored in EMU, so subsequent zoom
 *   only rescales the stored EMU and never recaptures. Authored-only metrics
 *   are never substituted.
 * - Only tagged `twoCellAnchor` + `editAs="oneCell"` anchors are captured.
 *   An ordinary sheet yields `undefined`: no reference allocation and no
 *   cell walk (only the anchor arrays are visited).
 * - A reference holds primitives only (anchor family, array index, the eight
 *   marker facts, EMU size). It survives structured clone to the worker and
 *   never retains anchors, shape graphs, size maps or geometry.
 * - Binding validates association with one projection (anchor counts,
 *   indices, eligibility, exact markers) and throws on mismatch. Bindings
 *   are keyed weakly by projection object, so viewers never share them.
 * - Lookup is O(1). Once bound, an eligible anchor never silently falls back
 *   to its current (edited) markers: a released reference, or anchor arrays
 *   replaced after binding, yield a non-renderable size; a non-positive
 *   captured size stays non-renderable even if edits make markers positive.
 *
 * Stateless capability limit: the public `renderViewport` has no edit
 * history. Without a bound reference it resolves tagged `editAs="oneCell"`
 * anchors from their current markers, which is the correct INITIAL display
 * only. Live band-edit retention is promised solely on the internal viewer
 * path that captures from the prepared projection and binds before edits.
 * No hidden global capture is ever taken from an arbitrary first paint.
 *
 * `editAs="absolute"` position freezing is outside this scope; its existing
 * marker-driven behaviour is unchanged and not claimed as implemented.
 */
import { EMU_PER_PX } from '@silurus/ooxml-core';
import type { Worksheet } from '../types.js';
import {
  resolveCellAnchorRect,
  type CellAnchorAxis,
  type CellAnchorSizeFacts,
  type RetainedAnchorExtent,
  type SheetAnchorRect,
} from './cell-anchor-geometry.js';

/** Anchor family of a compact entry: `ws.images` or `ws.shapeGroups`. */
export type InitialAnchorFamily = 'image' | 'shape';

/** One captured anchor. Primitive facts only (structured-clone safe). */
export interface InitialAnchorSizeEntry extends RetainedAnchorExtent {
  readonly family: InitialAnchorFamily;
  /** Index into the family's anchor array (acquisition order). */
  readonly index: number;
  readonly fromCol: number;
  readonly fromColOff: number;
  readonly fromRow: number;
  readonly fromRowOff: number;
  readonly toCol: number;
  readonly toColOff: number;
  readonly toRow: number;
  readonly toRowOff: number;
  /** Prepared initial display size in EMU; 0 when that rect was not renderable. */
  readonly cx: number;
  readonly cy: number;
}

/** Compact per-sheet prepared-initial reference (serializable). */
export interface InitialAnchorSizeReference {
  readonly version: 1;
  readonly imageCount: number;
  readonly shapeGroupCount: number;
  readonly entries: readonly InitialAnchorSizeEntry[];
}

interface CaptureAxes {
  readonly col: CellAnchorAxis;
  readonly row: CellAnchorAxis;
}

/** The existing GridGeometry surface used by capture
 * (`axesAtScale(initialScale)`). */
export interface InitialAnchorGeometry {
  axesAtScale(scale: number): CaptureAxes;
}

const MARKER_KEYS = [
  'fromCol', 'fromColOff', 'fromRow', 'fromRowOff',
  'toCol', 'toColOff', 'toRow', 'toRowOff',
] as const;

const NO_ANCHORS: readonly CellAnchorSizeFacts[] = Object.freeze([]);
const NON_RENDERABLE: RetainedAnchorExtent = Object.freeze({ cx: 0, cy: 0 });

interface InitialAnchorBinding {
  readonly reference: InitialAnchorSizeReference;
  /** Array identities at bind time; the projection already owns them. */
  readonly images: unknown;
  readonly shapeGroups: unknown;
  readonly sizes: WeakMap<object, RetainedAnchorExtent>;
}

const bindings = new WeakMap<Worksheet, InitialAnchorBinding>();
const releasedReferences = new WeakSet<InitialAnchorSizeReference>();

/** True only for anchors whose prepared initial size survives band edits. */
export function retainsInitialAnchorSize(anchor: CellAnchorSizeFacts): boolean {
  return anchor.anchorTag === 'twoCellAnchor' && anchor.editAs === 'oneCell';
}

/** Last one-based worksheet row whose prepared height determines an eligible
 * anchor's initial rect, or `undefined` when no anchor retains its initial
 * size. A marker edge is the top of its zero-based row plus `rowOff` EMU
 * (ECMA-376 Part 1 §20.5.2.33), so rows through each marker row are needed;
 * derived from the from/to markers only, never from anchor kind or size. */
export function initialAnchorRowCoverage(worksheet: Worksheet): number | undefined {
  let last: number | undefined;
  const families: ReadonlyArray<readonly CellAnchorSizeFacts[] | undefined> = [
    worksheet.images, worksheet.shapeGroups,
  ];
  for (const family of families) {
    for (const anchor of family ?? []) {
      if (!retainsInitialAnchorSize(anchor)) continue;
      const markers = anchor as CellAnchorSizeFacts & { readonly fromRow?: number; readonly toRow?: number };
      let row = 0;
      for (const marker of [markers.fromRow, markers.toRow]) {
        // Zero-based marker row -> one-based worksheet row.
        if (typeof marker === 'number' && Number.isFinite(marker) && marker >= 0) {
          row = Math.max(row, Math.floor(marker) + 1);
        }
      }
      last = Math.max(last ?? 0, row);
    }
  }
  return last;
}

function emuOf(px: number): number {
  const emu = px * EMU_PER_PX;
  return Number.isFinite(emu) && emu > 0 ? emu : 0;
}

/** A structured-cloned reference can reuse a worker projection, but a changed
 * payload cannot bypass validation at the same revision. Comparing only these
 * compact facts costs O(retained anchors), with no allocation or cell walk. */
export function sameInitialAnchorSizeReference(
  left: InitialAnchorSizeReference | undefined,
  right: InitialAnchorSizeReference | undefined,
): boolean {
  if (left === right) return true;
  if (!left || !right || left.version !== right.version
    || left.imageCount !== right.imageCount || left.shapeGroupCount !== right.shapeGroupCount
    || !Array.isArray(left.entries) || !Array.isArray(right.entries)
    || left.entries.length !== right.entries.length) return false;
  for (let index = 0; index < left.entries.length; index++) {
    const a = left.entries[index];
    const b = right.entries[index];
    if (!a || !b || a.family !== b.family || a.index !== b.index
      || !Object.is(a.cx, b.cx) || !Object.is(a.cy, b.cy)) return false;
    for (const key of MARKER_KEYS) if (!Object.is(a[key], b[key])) return false;
  }
  return true;
}

function mismatch(reason: string): never {
  throw new Error(
    `XLSX initial anchor size reference does not match this worksheet projection: ${reason}`,
  );
}

/**
 * Capture prepared initial sizes from `prepared` (host fonts bound, automatic
 * row heights prepared, no manual band edit yet). `geometry` must be that
 * worksheet's existing GridGeometry. Returns `undefined` when the sheet has
 * no eligible anchor.
 *
 * `initialScale` is the scale of the prepared first display (library policy).
 * The rect is resolved at that scale, where the grid rounds each band's
 * scaled pixels, then divided by it and stored in EMU (non-finite or
 * non-positive => 0, non-renderable). Later zoom only rescales the stored EMU.
 */
export function captureInitialAnchorSizes(
  prepared: Worksheet,
  geometry: InitialAnchorGeometry,
  initialScale = 1,
): InitialAnchorSizeReference | undefined {
  if (!Number.isFinite(initialScale) || initialScale <= 0) {
    throw new RangeError(
      `XLSX initial anchor size capture requires a finite positive scale: ${initialScale}`,
    );
  }
  const images: readonly CellAnchorSizeFacts[] = prepared.images ?? NO_ANCHORS;
  const shapeGroups: readonly CellAnchorSizeFacts[] = prepared.shapeGroups ?? NO_ANCHORS;
  let entries: InitialAnchorSizeEntry[] | undefined;
  let axes: CaptureAxes | undefined;
  for (let pass = 0; pass < 2; pass++) {
    const family: InitialAnchorFamily = pass === 0 ? 'image' : 'shape';
    const anchors = pass === 0 ? images : shapeGroups;
    for (let index = 0; index < anchors.length; index++) {
      const anchor = anchors[index];
      if (!retainsInitialAnchorSize(anchor)) continue;
      axes ??= geometry.axesAtScale(initialScale);
      // No retained input: a tagged anchor's initial rect is from/to.
      const rect = resolveCellAnchorRect(anchor, axes.col, axes.row, initialScale);
      (entries ??= []).push(Object.freeze({
        family,
        index,
        fromCol: anchor.fromCol,
        fromColOff: anchor.fromColOff,
        fromRow: anchor.fromRow,
        fromRowOff: anchor.fromRowOff,
        toCol: anchor.toCol,
        toColOff: anchor.toColOff,
        toRow: anchor.toRow,
        toRowOff: anchor.toRowOff,
        cx: emuOf(rect.width / initialScale),
        cy: emuOf(rect.height / initialScale),
      }));
    }
  }
  if (!entries) return undefined;
  return Object.freeze({
    version: 1,
    imageCount: images.length,
    shapeGroupCount: shapeGroups.length,
    entries: Object.freeze(entries),
  });
}

/**
 * Bind `reference` to one viewer-owned or worker render-local projection
 * BEFORE it is painted or edited. Validates the association and throws on
 * mismatch (never a silent fallback). Re-binding the same reference is a
 * no-op; a projection cannot be rebound to a different reference.
 */
export function bindInitialAnchorSizes(
  projection: Worksheet,
  reference: InitialAnchorSizeReference,
): void {
  if (releasedReferences.has(reference)) mismatch('the reference was released');
  const existing = bindings.get(projection);
  if (existing) {
    if (existing.reference === reference && projection.images === existing.images
      && projection.shapeGroups === existing.shapeGroups) return;
    mismatch('the projection is already bound to another reference');
  }
  if (reference.version !== 1 || !Array.isArray(reference.entries)) {
    mismatch('unsupported reference shape');
  }
  const images: readonly CellAnchorSizeFacts[] = projection.images ?? NO_ANCHORS;
  const shapeGroups: readonly CellAnchorSizeFacts[] = projection.shapeGroups ?? NO_ANCHORS;
  if (images.length !== reference.imageCount || shapeGroups.length !== reference.shapeGroupCount) {
    mismatch('anchor counts differ');
  }
  const sizes = new WeakMap<object, RetainedAnchorExtent>();
  for (const entry of reference.entries) {
    const anchors = entry.family === 'image' ? images
      : entry.family === 'shape' ? shapeGroups
        : NO_ANCHORS;
    const anchor = Number.isInteger(entry.index) ? anchors[entry.index] : undefined;
    if (!anchor || !retainsInitialAnchorSize(anchor) || sizes.has(anchor)) {
      mismatch(`${entry.family}[${entry.index}] is not an eligible anchor`);
    }
    for (const key of MARKER_KEYS) {
      if (!Object.is(anchor[key], entry[key])) {
        mismatch(`${entry.family}[${entry.index}].${key} differs`);
      }
    }
    if (!Number.isFinite(entry.cx) || !Number.isFinite(entry.cy) || entry.cx < 0 || entry.cy < 0) {
      mismatch(`${entry.family}[${entry.index}] has an invalid EMU size`);
    }
    sizes.set(anchor, entry);
  }
  let eligible = 0;
  for (const anchor of images) if (retainsInitialAnchorSize(anchor)) eligible++;
  for (const anchor of shapeGroups) if (retainsInitialAnchorSize(anchor)) eligible++;
  if (eligible !== reference.entries.length) {
    mismatch('an eligible anchor has no captured initial size');
  }
  // Capture already freezes its result. Seal the worker-owned structured clone
  // too, so an identical-reference cache shortcut cannot observe in-place edits.
  for (const entry of reference.entries) Object.freeze(entry);
  Object.freeze(reference.entries);
  Object.freeze(reference);
  bindings.set(projection, {
    reference,
    images: projection.images,
    shapeGroups: projection.shapeGroups,
    sizes,
  });
}

/**
 * Prepared initial size for `anchor` on `projection`, or `undefined` when
 * the anchor is not retained or the projection was never bound (stateless,
 * initial-display-only semantics). Bound but released/replaced anchors
 * resolve to a non-renderable size instead of their edited markers.
 */
export function lookupInitialAnchorSize(
  projection: Worksheet,
  anchor: CellAnchorSizeFacts,
): RetainedAnchorExtent | undefined {
  if (!retainsInitialAnchorSize(anchor)) return undefined;
  const binding = bindings.get(projection);
  if (!binding) return undefined;
  if (releasedReferences.has(binding.reference)) return NON_RENDERABLE;
  if (projection.images !== binding.images || projection.shapeGroups !== binding.shapeGroups) {
    return NON_RENDERABLE;
  }
  // Once this projection is bound, an unknown eligible object has no initial
  // size. Ordinary charts are untagged and exit before this lookup.
  return binding.sizes.get(anchor) ?? NON_RENDERABLE;
}

/** Display rectangle of `anchor` on `projection`: the single resolver applied
 * with that projection's bound prepared-initial size (if any). Paint, shape
 * groups, hit/outline projection and culling/decode sizing all call this so
 * they always agree on the same rectangle. */
export function resolveWorksheetAnchorRect(
  projection: Worksheet,
  anchor: CellAnchorSizeFacts,
  colAxis: CellAnchorAxis,
  rowAxis: CellAnchorAxis,
  scale: number,
): SheetAnchorRect {
  return resolveCellAnchorRect(
    anchor, colAxis, rowAxis, scale, lookupInitialAnchorSize(projection, anchor),
  );
}

/** Mark a reference released (workbook reset / viewer destroy). Projections
 * still bound to it paint its retained anchors as non-renderable rather than
 * reviving edited markers; it can no longer be bound. */
export function releaseInitialAnchorSizeReference(reference: InitialAnchorSizeReference): void {
  releasedReferences.add(reference);
}
