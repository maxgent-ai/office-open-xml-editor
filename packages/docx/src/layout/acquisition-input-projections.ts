import type { NumberingInfo } from '../types.js';
import type { ParagraphAcquisitionInput } from './text.js';
import type { ParagraphLayoutSource } from './text.js';
import type { TableSourceAcquisitionInput, TableLayoutSource } from './table-source-acquisition.js';
import type {
  DeepReadonly,
  NumberingMarkerShapeInput,
  SourceRef,
  TableColumnLayoutInput,
  TableFormatInput,
  TablePreferredWidthConstraint,
} from './types.js';

/** Parser-owned fact projections needed by otherwise parser-independent body
 * acquisition. Layout consumes this required capability record instead of
 * importing parser-model implementation or private wire fields. */
export interface BodyAcquisitionInputProjections {
  readonly numberingMarkerShapeInput: (
    numbering: NumberingInfo,
    fallbackFontSizePt: number,
  ) => NumberingMarkerShapeInput;
  readonly paragraphMarkShapeInput: (
    paragraph: ParagraphLayoutSource,
  ) => NumberingMarkerShapeInput | undefined;
  readonly tableFormatInput: (
    table: TableLayoutSource,
  ) => TableFormatInput;
  readonly tableSourceAcquisitionInput: (table: TableLayoutSource) => TableSourceAcquisitionInput;
  /** Effective whole-table preference, including the first-row exception. */
  readonly effectiveTablePreferredWidthPt: (
    table: TableLayoutSource,
    availableWidthPt: number,
  ) => number | null;
  readonly tableColumnLayoutInput: (
    table: TableLayoutSource,
    availableWidthPt: number,
    intrinsicWidths: (
      cell: TableLayoutSource['rows'][number]['cells'][number],
      preferredWidth: TablePreferredWidthConstraint | null,
    ) => Readonly<{ minWidthPt: number; maxWidthPt: number }>,
    maximumWidthPt?: number | null,
  ) => TableColumnLayoutInput;
  readonly tableParticipatesInOrdinaryFlow: (
    table: TableLayoutSource,
  ) => boolean;
  readonly paragraphAcquisitionInput: (
    paragraph: ParagraphLayoutSource,
    source: SourceRef,
  ) => ParagraphAcquisitionInput;
}
