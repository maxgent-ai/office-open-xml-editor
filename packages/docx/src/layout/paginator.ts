import {
  sectionContentEndBlockPt,
  sectionContentStartBlockPt,
  type PageFlowSectionContext,
} from './context.js';
import {
  createSectionRegionCoordinateSpace,
  sectionWritingMode,
  transformRect,
  uprightPhysicalExtent,
} from './coordinate-space.js';
import type { PaintNode } from './types.js';

export type PageAdvanceReason =
  | 'overflow'
  | 'explicit-break'
  | 'page-break-before'
  | 'section-break'
  | 'parity';

export type SectionStartType =
  | 'continuous'
  | 'nextColumn'
  | 'nextPage'
  | 'oddPage'
  | 'evenPage';

export type AuthoredBreak =
  | 'column'
  | 'page'
  | 'pageBreakBefore'
  | 'lastRenderedPageBreak';

export type PhysicalPageParity = 'odd' | 'even';

export class UnsupportedPageFlowTransitionError extends Error {
  readonly code = 'NEXT_COLUMN_DESTINATION_UNAVAILABLE' as const;

  constructor(
    readonly outgoingColumnIndex: number,
    readonly outgoingColumnCount: number,
    readonly incomingColumnCount: number,
    readonly reason:
      | 'page-extent'
      | 'writing-mode'
      | 'block-band'
      | 'grid'
      | 'physical-overlap'
      | 'physical-column',
  ) {
    super(
      'nextColumn requires a following column on the current page, '
      + `but column ${outgoingColumnIndex + 1} is unavailable `
      + `(outgoing columns: ${outgoingColumnCount}, incoming columns: ${incomingColumnCount}, `
      + `reason: ${reason})`,
    );
    this.name = 'UnsupportedPageFlowTransitionError';
  }
}

export interface SectionBoundaryOptions {
  readonly hasFootnoteReferenceOnCurrentPage?: boolean;
  readonly incomingPageContentStartBlockPt?: number;
  readonly incomingPageContentEndBlockPt?: number;
}

export interface PageFlowState {
  readonly pageIndex: number;
  readonly columnIndex: number;
  /** Whether this physical page already owns placed body content. */
  readonly pageHasContent: boolean;
  /** Page-absolute logical block coordinate (pt), independent of writing mode. */
  readonly cursorBlockPt: number;
  /** Logical block origin of the physical page's body content. */
  readonly pageContentStartBlockPt: number;
  readonly pageContentEndBlockPt: number;
  /** Logical block origin shared by every column in the active section region. */
  readonly regionStartBlockPt: number;
  /** Logical block end shared by every column in the active section region. */
  readonly regionEndBlockPt: number;
  /** Authored section-column indexes owned by the active region, in physical order. */
  readonly columnSubset: readonly number[];
  /** Deepest block edge reached by any completed/current column in the region. */
  readonly deepestColumnBlockPt: number;
  readonly section: PageFlowSectionContext;
}

export type PageFlowEvent =
  | Readonly<{
      type: 'place';
      node: PaintNode;
      blockStartPt: number;
      blockEndPt: number;
    }>
  | Readonly<{ type: 'next-column' }>
  | Readonly<{
      type: 'next-page';
      reason: PageAdvanceReason;
      pageIndex: number;
      sectionOccurrenceId: string;
      parityBlank: boolean;
    }>
  | Readonly<{
      type: 'begin-section';
      section: PageFlowSectionContext;
    }>
  | Readonly<{
      type: 'begin-section';
      placement: 'same-page-block' | 'same-page-column';
      section: PageFlowSectionContext;
      targetColumnOrdinal: number;
      columnSubset: readonly number[];
      outgoingColumnSubset?: readonly number[];
    }>;

export interface PageFlowTransition {
  readonly state: PageFlowState;
  readonly events: readonly PageFlowEvent[];
}

function columnPopulationOrder(
  section: PageFlowSectionContext,
  columnSubset: readonly number[],
): readonly number[] {
  return section.sectionBidi ? [...columnSubset].reverse() : [...columnSubset];
}

function followingColumnIndex(state: PageFlowState): number | undefined {
  const populationOrder = columnPopulationOrder(state.section, state.columnSubset);
  const currentOrdinal = populationOrder.indexOf(state.columnIndex);
  return populationOrder[currentOrdinal + 1];
}

export function createPageFlowState(
  section: PageFlowSectionContext,
  overrides: Partial<Omit<PageFlowState, 'section'>> = {},
): PageFlowState {
  const contentStart = sectionContentStartBlockPt(section);
  const contentEnd = sectionContentEndBlockPt(section);
  const pageContentStartBlockPt = overrides.pageContentStartBlockPt ?? contentStart;
  const pageContentEndBlockPt = overrides.pageContentEndBlockPt ?? contentEnd;
  const regionStartBlockPt = overrides.regionStartBlockPt ?? pageContentStartBlockPt;
  const regionEndBlockPt = overrides.regionEndBlockPt ?? pageContentEndBlockPt;
  const cursorBlockPt = overrides.cursorBlockPt ?? regionStartBlockPt;
  const deepestColumnBlockPt = overrides.deepestColumnBlockPt ?? cursorBlockPt;
  const pageIndex = overrides.pageIndex ?? 0;
  const columnSubset = Object.freeze([
    ...(overrides.columnSubset ?? section.columns.map((_, index) => index)),
  ]);
  const populationOrder = columnPopulationOrder(section, columnSubset);
  const columnIndex = overrides.columnIndex ?? populationOrder[0] ?? -1;
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Page index must be a non-negative integer');
  }
  if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= section.columns.length) {
    throw new RangeError('Column index must identify a column in the active section');
  }
  if (
    columnSubset.length === 0
    || columnSubset.some((index, position) => (
      !Number.isInteger(index)
      || index < 0
      || index >= section.columns.length
      || (position > 0 && index <= columnSubset[position - 1]!)
    ))
    || !columnSubset.includes(columnIndex)
  ) {
    throw new RangeError('Column subset must be ordered, unique, and contain the active column');
  }
  if (![
    pageContentStartBlockPt,
    pageContentEndBlockPt,
    regionStartBlockPt,
    regionEndBlockPt,
    cursorBlockPt,
    deepestColumnBlockPt,
  ].every(Number.isFinite)) {
    throw new RangeError('Page-flow cursors and bounds must be finite');
  }
  if (
    pageContentStartBlockPt > regionStartBlockPt
    || regionStartBlockPt > regionEndBlockPt
    || regionEndBlockPt > pageContentEndBlockPt
    || regionStartBlockPt > cursorBlockPt
    || cursorBlockPt > regionEndBlockPt
    || cursorBlockPt > deepestColumnBlockPt
  ) {
    throw new RangeError(
      'Page-flow bounds must contain the region and live cursor',
    );
  }
  return Object.freeze({
    pageIndex,
    columnIndex,
    pageHasContent: overrides.pageHasContent ?? false,
    cursorBlockPt,
    pageContentStartBlockPt,
    pageContentEndBlockPt,
    regionStartBlockPt,
    regionEndBlockPt,
    columnSubset,
    deepestColumnBlockPt,
    section,
  });
}

function transition(
  state: PageFlowState,
  events: readonly PageFlowEvent[],
): PageFlowTransition {
  return Object.freeze({
    state,
    events: Object.freeze(events.map((event) => Object.freeze({ ...event }))),
  });
}

export function placeFlowNode(
  state: PageFlowState,
  node: PaintNode,
  flowChargePt: number,
): PageFlowTransition {
  if (!Number.isFinite(flowChargePt) || flowChargePt < 0) {
    throw new RangeError('A flow node charge must be a finite non-negative value');
  }
  const blockStartPt = state.cursorBlockPt;
  const blockEndPt = blockStartPt + flowChargePt;
  return transition(Object.freeze({
    ...state,
    pageHasContent: true,
    cursorBlockPt: blockEndPt,
    deepestColumnBlockPt: Math.max(state.deepestColumnBlockPt, blockEndPt),
  }), [{
    type: 'place',
    node,
    blockStartPt,
    blockEndPt,
  }]);
}

export function advanceColumnOrPage(
  state: PageFlowState,
  reason: Extract<PageAdvanceReason, 'overflow' | 'explicit-break'>,
): PageFlowTransition {
  const deepestColumnBlockPt = Math.max(
    state.deepestColumnBlockPt,
    state.cursorBlockPt,
  );
  const nextColumnIndex = followingColumnIndex(state);
  if (nextColumnIndex !== undefined) {
    return transition(Object.freeze({
      ...state,
      columnIndex: nextColumnIndex,
      cursorBlockPt: state.regionStartBlockPt,
      deepestColumnBlockPt,
    }), [{ type: 'next-column' }]);
  }

  const pageIndex = state.pageIndex + 1;
  return transition(createPageFlowState(state.section, { pageIndex }), [{
    type: 'next-page',
    reason,
    pageIndex,
    sectionOccurrenceId: state.section.sectionOccurrenceId,
    parityBlank: false,
  }]);
}

interface NextColumnDestination {
  readonly targetColumnIndex: number;
  readonly targetColumnOrdinal: number;
  readonly columnSubset: readonly number[];
  readonly outgoingColumnSubset: readonly number[];
}

function sameGrid(
  left: PageFlowSectionContext['grid'],
  right: PageFlowSectionContext['grid'],
): boolean {
  return left.kind === right.kind
    && left.linePitchPt === right.linePitchPt
    && left.charSpacePt === right.charSpacePt;
}

function sameRect(
  left: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
  right: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
): boolean {
  return left.xPt === right.xPt
    && left.yPt === right.yPt
    && left.widthPt === right.widthPt
    && left.heightPt === right.heightPt;
}

function rectsOverlap(
  left: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
  right: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
): boolean {
  return left.xPt < right.xPt + right.widthPt
    && right.xPt < left.xPt + left.widthPt
    && left.yPt < right.yPt + right.heightPt
    && right.yPt < left.yPt + left.heightPt;
}

/** Resolve §17.18.77 against retained physical bands. Column indexes are not
 * transferable between section occurrences because §17.6.4 permits each
 * occurrence to author a different grid. */
function resolveNextColumnDestination(
  state: PageFlowState,
  section: PageFlowSectionContext,
  options: SectionBoundaryOptions,
): NextColumnDestination {
  const reject = (
    reason: UnsupportedPageFlowTransitionError['reason'],
  ): never => {
    throw new UnsupportedPageFlowTransitionError(
      state.columnIndex,
      state.section.columns.length,
      section.columns.length,
      reason,
    );
  };
  const outgoingWritingMode = sectionWritingMode(state.section);
  const incomingWritingMode = sectionWritingMode(section);
  if (outgoingWritingMode !== incomingWritingMode) reject('writing-mode');
  const outgoingPage = uprightPhysicalExtent({
    widthPt: state.section.geometry.pageWidth,
    heightPt: state.section.geometry.pageHeight,
  }, outgoingWritingMode);
  const incomingPage = uprightPhysicalExtent({
    widthPt: section.geometry.pageWidth,
    heightPt: section.geometry.pageHeight,
  }, incomingWritingMode);
  if (
    outgoingPage.widthPt !== incomingPage.widthPt
    || outgoingPage.heightPt !== incomingPage.heightPt
  ) reject('page-extent');
  const incomingContentStart = options.incomingPageContentStartBlockPt
    ?? sectionContentStartBlockPt(section);
  const incomingContentEnd = options.incomingPageContentEndBlockPt
    ?? sectionContentEndBlockPt(section);
  if (
    incomingContentStart !== state.pageContentStartBlockPt
    || incomingContentEnd !== state.pageContentEndBlockPt
  ) reject('block-band');
  if (!sameGrid(state.section.grid, section.grid)) reject('grid');

  const outgoingPopulation = columnPopulationOrder(state.section, state.columnSubset);
  const currentOrdinal = outgoingPopulation.indexOf(state.columnIndex);
  const successorIndex = outgoingPopulation[currentOrdinal + 1];
  if (successorIndex === undefined) {
    throw new Error('nextColumn destination resolution requires a same-page successor');
  }
  const coordinateSpace = createSectionRegionCoordinateSpace(
    outgoingWritingMode,
    outgoingPage,
  );
  const outgoingColumn = state.section.columns[successorIndex]!;
  const outgoingPhysicalBand = transformRect(coordinateSpace.logicalToPhysical, {
    xPt: outgoingColumn.xPt,
    yPt: state.regionStartBlockPt,
    widthPt: outgoingColumn.wPt,
    heightPt: state.regionEndBlockPt - state.regionStartBlockPt,
  });
  const targetColumnIndex = section.columns.findIndex((column) => (
    sameRect(outgoingPhysicalBand, transformRect(coordinateSpace.logicalToPhysical, {
      xPt: column.xPt,
      yPt: state.regionStartBlockPt,
      widthPt: column.wPt,
      heightPt: state.regionEndBlockPt - state.regionStartBlockPt,
    }))
  ));
  if (targetColumnIndex < 0) reject('physical-column');
  const incomingPopulation = columnPopulationOrder(
    section,
    section.columns.map((_, index) => index),
  );
  const targetColumnOrdinal = incomingPopulation.indexOf(targetColumnIndex);
  if (targetColumnOrdinal < 0) reject('physical-column');
  const columnSubset = Object.freeze(
    incomingPopulation.slice(targetColumnOrdinal).sort((left, right) => left - right),
  );
  const outgoingColumnSubset = Object.freeze(
    outgoingPopulation.slice(0, currentOrdinal + 1).sort((left, right) => left - right),
  );
  const physicalBand = (
    owner: PageFlowSectionContext,
    columnIndex: number,
  ) => {
    const column = owner.columns[columnIndex]!;
    return transformRect(coordinateSpace.logicalToPhysical, {
      xPt: column.xPt,
      yPt: state.regionStartBlockPt,
      widthPt: column.wPt,
      heightPt: state.regionEndBlockPt - state.regionStartBlockPt,
    });
  };
  const outgoingOwnedBands = outgoingColumnSubset.map((columnIndex) =>
    physicalBand(state.section, columnIndex));
  if (columnSubset.some((columnIndex) => {
    const incomingBand = physicalBand(section, columnIndex);
    return outgoingOwnedBands.some((outgoingBand) => rectsOverlap(outgoingBand, incomingBand));
  })) {
    reject('physical-overlap');
  }
  return Object.freeze({
    targetColumnIndex,
    targetColumnOrdinal,
    columnSubset,
    outgoingColumnSubset,
  });
}

export function advanceToPage(
  state: PageFlowState,
  section: PageFlowSectionContext,
  reason: Extract<
    PageAdvanceReason,
    'overflow' | 'explicit-break' | 'page-break-before' | 'section-break'
  >,
): PageFlowTransition {
  const pageIndex = state.pageIndex + 1;
  return transition(createPageFlowState(section, { pageIndex }), [{
    type: 'next-page',
    reason,
    pageIndex,
    sectionOccurrenceId: section.sectionOccurrenceId,
    parityBlank: false,
  }]);
}

function matchesPhysicalPageParity(
  pageIndex: number,
  parity: PhysicalPageParity,
): boolean {
  const isOddPhysicalPage = pageIndex % 2 === 0;
  return parity === 'odd' ? isOddPhysicalPage : !isOddPhysicalPage;
}

function advanceToPageWithParity(
  state: PageFlowState,
  section: PageFlowSectionContext,
  reason: Extract<PageAdvanceReason, 'explicit-break' | 'section-break'>,
  parity?: PhysicalPageParity,
): PageFlowTransition {
  let pageIndex = state.pageIndex + 1;
  const events: PageFlowEvent[] = [];
  if (parity !== undefined && !matchesPhysicalPageParity(pageIndex, parity)) {
    events.push({
      type: 'next-page',
      reason: 'parity',
      pageIndex,
      sectionOccurrenceId: state.section.sectionOccurrenceId,
      parityBlank: true,
    });
    pageIndex += 1;
  }
  events.push({
    type: 'next-page',
    reason,
    pageIndex,
    sectionOccurrenceId: section.sectionOccurrenceId,
    parityBlank: false,
  });
  return transition(createPageFlowState(section, { pageIndex }), events);
}

export function applyAuthoredBreak(
  state: PageFlowState,
  authoredBreak: AuthoredBreak,
  parity?: PhysicalPageParity,
): PageFlowTransition {
  if (authoredBreak === 'lastRenderedPageBreak') {
    // lastRenderedPageBreak is a cached result from a previous layout producer,
    // not document intent. Mixing it with fresh pagination double-applies breaks.
    return transition(state, []);
  }
  if (authoredBreak === 'column') {
    return advanceColumnOrPage(state, 'explicit-break');
  }
  if (
    authoredBreak === 'pageBreakBefore'
    && !state.pageHasContent
    && state.columnIndex === columnPopulationOrder(state.section, state.columnSubset)[0]
    && state.cursorBlockPt === state.pageContentStartBlockPt
  ) {
    // §17.3.1.23 requires the paragraph to begin on a new page. A paragraph
    // already at the start of an otherwise empty page satisfies that condition.
    return transition(state, []);
  }
  return authoredBreak === 'page'
    ? advanceToPageWithParity(state, state.section, 'explicit-break', parity)
    : advanceToPage(state, state.section, 'page-break-before');
}

export function beginSection(
  state: PageFlowState,
  section: PageFlowSectionContext,
  startType: SectionStartType,
  options: SectionBoundaryOptions = {},
): PageFlowTransition {
  if (startType === 'continuous' && !options.hasFootnoteReferenceOnCurrentPage) {
    // §17.6.4: a section following newspaper columns begins below the deepest
    // column, not merely below the last column visited by source order.
    const regionTop = state.section.columns.length > 1
      ? Math.max(state.cursorBlockPt, state.deepestColumnBlockPt)
      : state.cursorBlockPt;
    return transition(createPageFlowState(section, {
      pageIndex: state.pageIndex,
      pageContentStartBlockPt: state.pageContentStartBlockPt,
      pageContentEndBlockPt: state.pageContentEndBlockPt,
      cursorBlockPt: regionTop,
      regionStartBlockPt: regionTop,
      regionEndBlockPt: state.pageContentEndBlockPt,
      deepestColumnBlockPt: regionTop,
      pageHasContent: state.pageHasContent,
    }), [{
      type: 'begin-section',
      placement: 'same-page-block',
      section,
      targetColumnOrdinal: 0,
      columnSubset: section.columns.map((_, index) => index),
    }]);
  }

  if (startType === 'nextColumn') {
    if (followingColumnIndex(state) === undefined) {
      // §17.18.77 defines nextColumn as the following column on the page. Once
      // the active region has no such column (including a one-column section),
      // the next flow destination is the incoming section's first column on the
      // following page. This is the retained solver's total transition policy
      // for a valid break whose same-page destination does not exist.
      const nextPage = advanceToPage(state, section, 'section-break');
      return transition(nextPage.state, [
        ...nextPage.events,
        { type: 'begin-section', section },
      ]);
    }
    const destination = resolveNextColumnDestination(state, section, options);
    return transition(Object.freeze({
      ...state,
      columnIndex: destination.targetColumnIndex,
      columnSubset: destination.columnSubset,
      cursorBlockPt: state.regionStartBlockPt,
      deepestColumnBlockPt: Math.max(
        state.deepestColumnBlockPt,
        state.cursorBlockPt,
      ),
      section,
    }), [
      { type: 'next-column' },
      {
        type: 'begin-section',
        placement: 'same-page-column',
        section,
        targetColumnOrdinal: destination.targetColumnOrdinal,
        columnSubset: destination.columnSubset,
        outgoingColumnSubset: destination.outgoingColumnSubset,
      },
    ]);
  }

  if (startType === 'continuous') {
    // §17.18.77 requires the continuous section to begin on the following page
    // when a footnote reference on this page would otherwise cross the boundary.
    const nextPage = advanceToPage(state, section, 'section-break');
    return transition(nextPage.state, [
      ...nextPage.events,
      { type: 'begin-section', section },
    ]);
  }

  const parity = startType === 'oddPage'
    ? 'odd'
    : startType === 'evenPage'
      ? 'even'
      : undefined;
  const pageAdvance = advanceToPageWithParity(state, section, 'section-break', parity);
  return transition(pageAdvance.state, [
    ...pageAdvance.events,
    { type: 'begin-section', section },
  ]);
}
