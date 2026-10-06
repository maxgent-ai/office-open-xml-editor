import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  type NormalizedXlsxWorksheetPolicy,
} from '@silurus/ooxml-core/worker';
import { cappedAdd, utf8Bytes } from '@silurus/ooxml-core/internal/resource-measurement';
import type { ParsedWorkbook, Row, Styles, Worksheet } from './types.js';
import { bindWorksheetPolicy } from './worksheet-policy-context.js';
import {
  XLSX_MAX_MATERIALIZED_JSON_BYTES,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  measureWorksheet,
  worksheetLimitError,
} from './worksheet-resource-limits.js';
import { MAX_WORKSHEET_COL } from './internal/grid-geometry.js';

/** Per-load source selection for {@link XlsxSheetViewer.load}. XLSX remains the
 * default; delimited text is an intentionally small reuse of the XLSX sheet
 * renderer rather than a general-purpose tabular-data API. */
export type XlsxSheetLoadOptions =
  | Readonly<{ format?: 'xlsx' }>
  | Readonly<{
      format: 'csv' | 'tsv';
      /** Override the format's comma/tab default with one literal character. */
      delimiter?: string;
      /** Browser TextDecoder label. Defaults to UTF-8. */
      encoding?: string;
      /** Model name exposed by `sheetNames[0]`. Defaults to `Sheet1`. */
      sheetName?: string;
    }>
  | Readonly<{
      /** Generic delimited text, including `.txt`, `.dat`, and `.psv` files. */
      format: 'delimited-text';
      /** One literal field-separator character. */
      delimiter: string;
      /** Browser TextDecoder label. Defaults to UTF-8. */
      encoding?: string;
      /** Model name exposed by `sheetNames[0]`. Defaults to `Sheet1`. */
      sheetName?: string;
    }>;

type DelimitedTextLoadOptions = Exclude<
  XlsxSheetLoadOptions,
  Readonly<{ format?: 'xlsx' }>
>;

export interface ResolvedDelimitedTextOptions {
  readonly delimiter: string;
  readonly encoding: string;
  readonly sheetName: string;
}

const DELIMITED_TEXT_OPERATION = 'load-delimited-text';

/** UTF-8 bytes of the retained `'text'` value discriminator that every
 * non-empty delimited field carries in its Cell value. */
const DELIMITED_TEXT_CELL_DISCRIMINATOR_UTF8_BYTES = 4;

/** Hard input ceiling. Parsing can expand delimiters into a larger worksheet
 * model, so this is only the first of the existing worksheet admission gates. */
export const DELIMITED_TEXT_MAX_SOURCE_BYTES = XLSX_MAX_MATERIALIZED_JSON_BYTES;

export function assertDelimitedTextSourceBytes(observed: number): void {
  if (observed > DELIMITED_TEXT_MAX_SOURCE_BYTES) {
    throw worksheetLimitError(
      DELIMITED_TEXT_OPERATION,
      undefined,
      'delimited-text-source',
      'bytes',
      DELIMITED_TEXT_MAX_SOURCE_BYTES,
      observed,
    );
  }
}

const DEFAULT_STYLES: Styles = Object.freeze({
  fonts: [Object.freeze({
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: null,
    name: 'Calibri',
  })],
  fills: [Object.freeze({ patternType: 'none', fgColor: null, bgColor: null })],
  borders: [Object.freeze({ left: null, right: null, top: null, bottom: null })],
  cellXfs: [Object.freeze({
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: null,
    alignV: null,
    wrapText: false,
  })],
  numFmts: [],
  dxfs: [],
}) as Styles;

export function resolveDelimitedTextOptions(
  options: DelimitedTextLoadOptions,
): ResolvedDelimitedTextOptions {
  if (
    !options
    || (
      options.format !== 'csv'
      && options.format !== 'tsv'
      && options.format !== 'delimited-text'
    )
  ) {
    throw new TypeError("format must be 'csv', 'tsv', or 'delimited-text'");
  }
  const delimiter = options.delimiter
    ?? (options.format === 'tsv' ? '\t' : options.format === 'csv' ? ',' : undefined);
  if (delimiter === undefined) {
    throw new TypeError("delimiter is required for format 'delimited-text'");
  }
  if (delimiter.length !== 1) {
    throw new TypeError('delimiter must be exactly one character');
  }
  if (delimiter === '"' || delimiter === '\r' || delimiter === '\n') {
    throw new TypeError('delimiter cannot be a quote or record separator');
  }
  const encoding = options.encoding ?? 'utf-8';
  if (typeof encoding !== 'string' || encoding.trim() === '') {
    throw new TypeError('encoding must be a non-empty TextDecoder label');
  }
  const sheetName = options.sheetName ?? 'Sheet1';
  if (typeof sheetName !== 'string' || sheetName.trim() === '') {
    throw new TypeError('sheetName must be a non-empty string');
  }
  return Object.freeze({ delimiter, encoding, sheetName });
}

/** Convert one bounded delimited-text source into the exact XLSX renderer
 * model. Every non-empty field remains text; no locale-sensitive type or
 * formula inference is performed. */
export function parseDelimitedWorksheet(
  source: ArrayBuffer,
  options: ResolvedDelimitedTextOptions,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): Readonly<{ workbook: ParsedWorkbook; worksheet: Worksheet }> {
  assertDelimitedTextSourceBytes(source.byteLength);
  const limits = policy.worksheet;

  let text: string;
  try {
    text = new TextDecoder(options.encoding, { fatal: true }).decode(source);
  } catch (error) {
    throw new TypeError(
      `Delimited text could not be decoded as ${options.encoding}`,
      { cause: error },
    );
  }

  const rows: Row[] = [];
  let cells: Row['cells'] = [];
  let field = '';
  let fieldChunks: string[] | undefined;
  let rowIndex = 1;
  let columnIndex = 1;
  /* Independent pre-expansion parsed-field guard. It counts every logical
   * field, blanks included, against the same public maxCells budget, so a
   * delimiter-only source cannot expand without bound before the terminal
   * admission. This is distinct from the terminal Cell record count, which
   * only sees non-empty fields retained as Cell records. */
  let logicalCellCount = 0;
  let retainedUtf8Bytes = 0;
  let quoted = false;
  let afterQuote = false;

  const appendField = (value: string): void => {
    field += value;
    // Keep one pathological field linear-time without retaining one array item
    // per character. The source and retained-text ceilings remain authoritative.
    if (field.length >= 4096) {
      (fieldChunks ??= []).push(field);
      field = '';
    }
  };

  const finishField = (): void => {
    logicalCellCount = cappedAdd(logicalCellCount, 1, limits.maxCells);
    if (logicalCellCount > limits.maxCells) {
      throw worksheetLimitError(
        DELIMITED_TEXT_OPERATION,
        undefined,
        'worksheet-model',
        'cells',
        limits.maxCells,
        logicalCellCount,
      );
    }
    if (columnIndex > MAX_WORKSHEET_COL) {
      throw new RangeError(
        `Delimited text row ${rowIndex} has more than ${MAX_WORKSHEET_COL} columns`,
      );
    }
    const fieldText = fieldChunks ? fieldChunks.join('') + field : field;
    if (fieldText !== '') {
      const maxOwned = limits.maxOwnedUtf8Bytes;
      retainedUtf8Bytes = cappedAdd(
        cappedAdd(retainedUtf8Bytes, DELIMITED_TEXT_CELL_DISCRIMINATOR_UTF8_BYTES, maxOwned),
        utf8Bytes(fieldText, maxOwned + 1),
        maxOwned,
      );
      if (retainedUtf8Bytes > maxOwned) {
        throw worksheetLimitError(
          DELIMITED_TEXT_OPERATION,
          undefined,
          'worksheet-cell-content',
          'owned-utf8-bytes',
          maxOwned,
          retainedUtf8Bytes,
        );
      }
      cells.push({
        row: rowIndex,
        col: columnIndex,
        value: { type: 'text', text: fieldText },
      });
    }
    field = '';
    fieldChunks = undefined;
    columnIndex++;
    afterQuote = false;
  };

  const finishRow = (): void => {
    if (rowIndex > limits.maxRows) {
      throw worksheetLimitError(
        DELIMITED_TEXT_OPERATION,
        undefined,
        'worksheet-model',
        'rows',
        limits.maxRows,
        rowIndex,
      );
    }
    rows.push({ index: rowIndex, height: null, cells });
    cells = [];
    rowIndex++;
    columnIndex = 1;
  };

  for (let index = 0; index < text.length; index++) {
    const character = text[index]!;
    if (quoted) {
      if (character !== '"') {
        if (character === '\r') {
          appendField('\n');
          if (text[index + 1] === '\n') index++;
        } else {
          appendField(character);
        }
        continue;
      }
      if (text[index + 1] === '"') {
        appendField('"');
        index++;
      } else {
        quoted = false;
        afterQuote = true;
      }
      continue;
    }

    if (afterQuote) {
      if (character !== options.delimiter && character !== '\r' && character !== '\n') {
        throw new SyntaxError(
          `Unexpected character after closing quote at row ${rowIndex}, column ${columnIndex}`,
        );
      }
    } else if (character === '"' && field === '' && fieldChunks === undefined) {
      quoted = true;
      continue;
    }

    if (character === options.delimiter) {
      finishField();
      continue;
    }
    if (character === '\r' || character === '\n') {
      finishField();
      finishRow();
      if (character === '\r' && text[index + 1] === '\n') index++;
      continue;
    }
    appendField(character);
  }

  if (quoted) {
    throw new SyntaxError(
      `Unterminated quoted field at row ${rowIndex}, column ${columnIndex}`,
    );
  }
  const endsWithRecordBreak = text.endsWith('\n') || text.endsWith('\r');
  if (text.length > 0 && !endsWithRecordBreak) {
    finishField();
    finishRow();
  }

  const worksheet: Worksheet = {
    name: options.sheetName,
    rows,
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  };
  const measured = measureWorksheet(worksheet, policy);
  assertWorksheetModelUsage(measured, DELIMITED_TEXT_OPERATION, undefined, undefined, policy);
  assertWorksheetJsonBytes(
    measured.jsonBytes,
    DELIMITED_TEXT_OPERATION,
    undefined,
    undefined,
    policy,
  );
  bindWorksheetPolicy(worksheet, policy);

  const workbook: ParsedWorkbook = {
    workbook: {
      sheets: [{ name: options.sheetName, sheetId: 1, rId: 'rId1' }],
    },
    styles: DEFAULT_STYLES,
    sharedStrings: [],
  };
  return { workbook, worksheet };
}
