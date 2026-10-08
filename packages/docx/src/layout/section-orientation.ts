import type { DocxDocumentModel, SectionProps } from '../types.js';
import type { NativeSectionFlow } from './types.js';
import {
  isVerticalSectionDirection,
  logicalSectionGeometry,
  physicalSectionGeometry,
} from './context.js';

/** True when a section flows vertically. The Transitional ST_TextDirection
 * values `tbRl`, `tbRlV`, `tbLrV`, and `btLr` share the rotated page frame (a
 * native BtoT section is nominally `btLr`; its own frame is selected by its
 * canonical `nativeSectionFlow`). */
export function isVerticalSection(section: SectionProps): boolean {
  return isVerticalTextDirection(section.textDirection);
}

/** Raw-token predicate for section-region carriers that do not own a complete
 * SectionProps value. Unknown and absent values remain horizontal. */
export function isVerticalTextDirection(
  textDirection: string | null | undefined,
): boolean {
  return typeof textDirection === 'string' && isVerticalSectionDirection(textDirection);
}

/** `btLr` uses the rotated page frame without upright-CJK counter-rotation. */
export function isAllRotatedVerticalTextDirection(
  textDirection: string | null | undefined,
): boolean {
  return textDirection === 'btLr';
}

/** Map physical vertical-section geometry into the logical horizontal frame
 * used by retained acquisition. ECMA-376 §17.6.11 defines the page margins. */
export function verticalLayoutSection(
  physical: SectionProps,
  nativeSectionFlow?: NativeSectionFlow,
): SectionProps {
  return {
    ...physical,
    ...logicalSectionGeometry(physical, nativeSectionFlow),
  };
}

/** Replace only the body-level section with its logical vertical frame.
 * Per-body section-break geometry remains physical until the paginator selects
 * each section's own frame. Horizontal documents preserve object identity. */
export function verticalLayoutDoc(document: DocxDocumentModel): DocxDocumentModel {
  if (!isVerticalSection(document.section)) return document;
  return { ...document, section: verticalLayoutSection(document.section) };
}

/** Map a logical vertical frame back to physical page geometry. */
export function physicalLayoutSection(
  logical: SectionProps,
  nativeSectionFlow?: NativeSectionFlow,
): SectionProps {
  return {
    ...logical,
    ...physicalSectionGeometry(logical, nativeSectionFlow),
  };
}
