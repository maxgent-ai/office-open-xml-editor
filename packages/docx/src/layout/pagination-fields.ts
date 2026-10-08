import { sourceOwnedTextPlacements, type SourceOwnedTextGeometry } from './text-source-ownership.js';
import type { NumberFormat } from '@silurus/ooxml-core';
import { convergeLayout, type LayoutIteration } from './convergence.js';
import { stableFingerprint } from './fingerprint.js';
import type { FlowFragment } from './flow-fragment.js';
import type { TableRowFragmentLayout } from './table-pagination.js';
import type {
  BodyElement,
  CellElement,
  DocNote,
  DocRun,
} from '../types.js';
import type { DocumentLayout } from './types.js';

export interface PaginationFieldPageContext {
  readonly pageIndex: number;
  readonly displayPageNumber: number;
  readonly pageNumberFormat: NumberFormat;
}

/**
 * Exact state read by the following field-acquisition pass. Body paint geometry
 * is deliberately absent: the deterministic transition consumes only total
 * page count (the array length) and these destination-page facts.
 */
export function paginationFieldPageContexts(
  layout: Pick<DocumentLayout, 'pages'>,
): readonly PaginationFieldPageContext[] {
  return Object.freeze(layout.pages.map((page) => Object.freeze({
    pageIndex: page.pageIndex,
    displayPageNumber: page.pageNumber.displayNumber,
    pageNumberFormat: page.pageNumber.format as NumberFormat,
  })));
}

// A resource-safety bound, not visual tuning. Sixteen exceeds the decimal digit
// transitions of any practical page count; convergeLayout's seen-set catches
// cycles earlier, while exhausting this limit hard-fails instead of returning
// stale field geometry.
const PAGINATION_FIELD_CONVERGENCE_LIMIT = 16;

/** Resolve PAGE/NUMPAGES-dependent pagination to a stable physical geometry. */
export function convergePaginationFields<T extends LayoutIteration>(
  acquire: (totalPagesHint: number) => T,
  limit = PAGINATION_FIELD_CONVERGENCE_LIMIT,
): T {
  const seed = acquire(1);
  return convergeLayout(seed, (current) => acquire(current.pageCount), limit);
}

/** The field dependencies whose measured result can change physical pagination. */
export function paginationFieldDependency(
  run: Extract<DocRun, { type: 'field' }>,
): 'page' | 'total-pages' | undefined {
  if (run.fieldType === 'page') return 'page';
  if (/numPages/i.test(run.fieldType) || /NUMPAGES/i.test(run.instruction)) return 'total-pages';
  return undefined;
}

function storyHasPaginationFields(elements: readonly (BodyElement | CellElement)[]): boolean {
  return elements.some((element) => {
    if (element.type === 'paragraph') {
      return element.runs.some((run) => {
        if (run.type !== 'field') return false;
        return paginationFieldDependency(run) !== undefined;
      });
    }
    if (element.type === 'table') {
      return element.rows.some((row) => row.cells.some((cell) =>
        storyHasPaginationFields(cell.content)));
    }
    return false;
  });
}

/**
 * Whether canonical body pagination has a field feedback edge that requires another pass.
 * Body PAGE needs its destination occurrence from the preceding pass. A footnote
 * PAGE likewise feeds its owning page's formatted number into note height and the
 * body reserve. Body and footnote NUMPAGES can feed measured width back into the
 * total page count. Header/footer reserve measurement is a separate story seam and
 * must not force an otherwise field-free body to repeat.
 */
export function paginatedFlowHasPaginationDependentFields(
  body: readonly BodyElement[],
  footnotes: readonly DocNote[] = [],
  additionalStories: readonly (readonly (BodyElement | CellElement)[])[] = [],
): boolean {
  return storyHasPaginationFields(body)
    || footnotes.some((note) => storyHasPaginationFields(note.content))
    || additionalStories.some((story) => storyHasPaginationFields(story));
}

/**
 * Field-free layout has no page-count feedback edge, so one acquisition is the
 * fixpoint. Keep the bounded convergence policy only when such an edge exists.
 */
export function resolvePaginationFieldLayout<T extends LayoutIteration>(
  acquire: (totalPagesHint: number) => T,
  hasPaginationFields: boolean,
): T {
  return hasPaginationFields ? convergePaginationFields(acquire) : acquire(1);
}

/** Project retained flow into convergence-relevant plain data. Field values are
 * retained because equal-width PAGE values can still belong to different pages. */
export function paginationFieldFlowGeometry(fragment: FlowFragment): unknown {
  if (fragment.kind === 'paragraph') {
    return {
      kind: fragment.kind,
      flowBounds: fragment.flowBounds,
      inkBounds: fragment.inkBounds,
      clipBounds: fragment.clipBounds,
      advancePt: fragment.advancePt,
      spacing: fragment.spacing,
      lines: fragment.lines.map((line) => ({
        range: line.range,
        bounds: line.bounds,
        wrapBounds: line.wrapBounds,
        wrapFragments: line.wrapFragments,
        baselinePt: line.baselinePt,
        advancePt: line.advancePt,
        placements: line.placements.flatMap<import('./types.js').ParagraphPlacement | SourceOwnedTextGeometry>(
          placement => placement.kind === 'text' ? sourceOwnedTextPlacements(placement) : [placement],
        ).map((placement) => ({
          kind: placement.kind,
          range: placement.range,
          bounds: placement.bounds,
          ...('advancePt' in placement ? { advancePt: placement.advancePt } : {}),
          ...(placement.kind === 'text' && placement.dependency
            ? {
                field: {
                  dependency: placement.dependency,
                  text: placement.text,
                  sourceRunIndex: placement.sourceRunIndex,
                },
              }
            : {}),
        })),
      })),
      drawings: fragment.drawings.map((drawing) => ({
        flowBounds: drawing.flowBounds,
        inkBounds: drawing.inkBounds,
        transform: drawing.transform,
        clip: drawing.clip,
      })),
      textBoxes: fragment.textBoxes.map((textBox) => ({
        flowBounds: textBox.flowBounds,
        inkBounds: textBox.inkBounds,
        advancePt: textBox.advancePt,
      })),
      exclusions: fragment.exclusions.map((exclusion) => ({
        wrap: exclusion.wrap,
        bounds: exclusion.bounds,
        polygon: exclusion.polygon,
      })),
    };
  }
  if (fragment.kind === 'table') {
    return {
      kind: fragment.kind,
      flowBounds: fragment.flowBounds,
      inkBounds: fragment.inkBounds,
      advancePt: fragment.advancePt,
      columnWidthsPt: fragment.columnWidthsPt,
      borders: fragment.borders,
      rows: fragment.rows.map((row) => ({
        ...('occurrenceId' in row ? {
          logicalRowIndex: (row as TableRowFragmentLayout).logicalRowIndex,
          fragmentIndex: (row as TableRowFragmentLayout).fragmentIndex,
          ownership: (row as TableRowFragmentLayout).ownership,
          occurrenceId: (row as TableRowFragmentLayout).occurrenceId,
          physicalPageIndex: (row as TableRowFragmentLayout).physicalPageIndex,
          displayPageNumber: (row as TableRowFragmentLayout).displayPageNumber,
        } : {}),
        flowBounds: row.flowBounds,
        contentHeightPt: row.contentHeightPt,
        cells: row.cells.map((cell) => ({
          flowBounds: cell.flowBounds,
          contentBounds: cell.contentBounds,
          verticalMerge: cell.verticalMerge,
          blocks: cell.blocks.map((block) => paginationFieldFlowGeometry(block.layout)),
        })),
      })),
    };
  }
  throw new Error('Unsupported retained flow fragment');
}

function definedRuntimeGeometry(value: unknown): unknown {
  if (Array.isArray(value)) return value.map((entry) => definedRuntimeGeometry(entry));
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .filter(([, entry]) => entry !== undefined)
        .map(([key, entry]) => [key, definedRuntimeGeometry(entry)]),
    );
  }
  return value;
}

/**
 * Fingerprint pagination geometry after omitting optional runtime placement
 * facts whose absence may be represented by a missing key or `undefined`.
 * This normalization is local to field convergence; other layout contracts keep
 * rejecting undefined data through the ordinary fingerprint boundary.
 */
export function paginationFieldGeometryFingerprint(value: unknown): string {
  return stableFingerprint(
    'pagination-field-geometry',
    definedRuntimeGeometry(value),
  );
}
