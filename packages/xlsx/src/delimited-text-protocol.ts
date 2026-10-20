import type { NormalizedXlsxWorksheetPolicy } from './internal/worksheet-policy.js';
import type {
  WorkerRendererDescriptors,
  WorkerErrorPayload,
} from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook } from './types.js';
import type { ResolvedDelimitedTextOptions } from './delimited-text.js';

export type DelimitedTextParseRequest = {
  readonly type: 'parseDelimitedText';
  readonly id: number;
  readonly data: ArrayBuffer;
  readonly options: ResolvedDelimitedTextOptions;
  readonly worksheetPolicy?: NormalizedXlsxWorksheetPolicy;
  readonly useGoogleFonts?: boolean;
  /** Caller-owned load-time cell-font language gate, shared with main layout. */
  readonly japaneseCellFonts?: boolean;
  readonly cjkFallback?: import('@silurus/ooxml-core').CjkLang;
  readonly renderers?: WorkerRendererDescriptors;
};

export type DelimitedTextParseResponse =
  | {
      readonly type: 'delimitedTextParsed';
      readonly id: number;
      readonly workbook: ParsedWorkbook;
      readonly worksheetJson: ArrayBuffer;
    }
  | ({ readonly type: 'error'; readonly id: number } & WorkerErrorPayload);
