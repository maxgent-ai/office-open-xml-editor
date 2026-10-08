import type { SectionLayoutContext } from '../layout-context.js';
import type { PageBorderEdge, PageBorders } from '../types.js';
import { retainedBorderTreatment } from './border-treatment.js';
import {
  createSectionRegionCoordinateSpace,
  sectionWritingMode,
  type PhysicalPageExtent,
} from './coordinate-space.js';
import { sectionBodyInsetPt } from './context.js';
import type {
  BorderSegment,
  PageBorderLayout,
} from './types.js';

function shownOnPage(
  pageBorders: PageBorders,
  firstSectionOwnedPage: boolean,
): boolean {
  switch (pageBorders.display) {
    case 'firstPage':
      return firstSectionOwnedPage;
    case 'notFirstPage':
      return !firstSectionOwnedPage;
    default:
      return true;
  }
}

function retainedColor(color: string | undefined): string {
  return color !== undefined && /^[0-9a-fA-F]{6}$/.test(color)
    ? `#${color}`
    : '#000000';
}

function retainedSpace(edge: PageBorderEdge | undefined): number {
  return edge !== undefined && Number.isFinite(edge.space) ? edge.space : 0;
}

function retainedSegment(
  edge: PageBorderEdge,
  edgeName: NonNullable<BorderSegment['edge']>,
  from: BorderSegment['from'],
  to: BorderSegment['to'],
): BorderSegment {
  const widthPt = Number.isFinite(edge.width) ? edge.width : 0.5;
  return Object.freeze({
    edge: edgeName,
    from: Object.freeze(from),
    to: Object.freeze(to),
    color: retainedColor(edge.color),
    widthPt,
    ...retainedBorderTreatment(edge.style, widthPt),
  });
}

/** Resolve one section-owned page decoration into immutable point-space paint
 * facts. ECMA-376 §17.6.10 makes `display` relative to the first physical page
 * owned by the section, including a parity blank. */
export function materializePageBorderLayout(
  pageBorders: PageBorders | null | undefined,
  section: SectionLayoutContext,
  physicalPage: PhysicalPageExtent,
  firstSectionOwnedPage: boolean,
): PageBorderLayout | null {
  if (!pageBorders || !shownOnPage(pageBorders, firstSectionOwnedPage)) return null;

  const { geometry } = section;
  const fromText = pageBorders.offsetFrom === 'text';
  const refLeftPt = fromText ? geometry.marginLeft : 0;
  const refRightPt = fromText
    ? geometry.pageWidth - geometry.marginRight
    : geometry.pageWidth;
  const refTopPt = fromText ? sectionBodyInsetPt(geometry.marginTop) : 0;
  const refBottomPt = fromText
    ? geometry.pageHeight - sectionBodyInsetPt(geometry.marginBottom)
    : geometry.pageHeight;
  const topY = refTopPt + retainedSpace(pageBorders.top);
  const bottomY = refBottomPt - retainedSpace(pageBorders.bottom);
  const leftX = refLeftPt + retainedSpace(pageBorders.left);
  const rightX = refRightPt - retainedSpace(pageBorders.right);
  const segments: BorderSegment[] = [];

  if (pageBorders.top) {
    segments.push(retainedSegment(
      pageBorders.top,
      'top',
      { xPt: leftX, yPt: topY },
      { xPt: rightX, yPt: topY },
    ));
  }
  if (pageBorders.bottom) {
    segments.push(retainedSegment(
      pageBorders.bottom,
      'bottom',
      { xPt: leftX, yPt: bottomY },
      { xPt: rightX, yPt: bottomY },
    ));
  }
  if (pageBorders.left) {
    segments.push(retainedSegment(
      pageBorders.left,
      'left',
      { xPt: leftX, yPt: topY },
      { xPt: leftX, yPt: bottomY },
    ));
  }
  if (pageBorders.right) {
    segments.push(retainedSegment(
      pageBorders.right,
      'right',
      { xPt: rightX, yPt: topY },
      { xPt: rightX, yPt: bottomY },
    ));
  }
  if (segments.length === 0) return null;

  const coordinateSpace = createSectionRegionCoordinateSpace(
    sectionWritingMode(section),
    physicalPage,
  );
  return Object.freeze({
    zOrder: pageBorders.zOrder === 'back' ? 'back' : 'front',
    logicalToPhysical: Object.freeze({ ...coordinateSpace.logicalToPhysical }),
    segments: Object.freeze(segments),
  });
}
