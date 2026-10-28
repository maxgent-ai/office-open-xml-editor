/**
 * XLSX font requests independent of font sources: the workbook's CJK fallback
 * language and the exact local Office face tuples the renderer preflights.
 * Font-source demand (`LoadOptions.fonts`) lives in the lazily loaded
 * `font-source-load.ts`.
 */
import {
  classifyCjkFont,
  cjkFallbackForText,
  findReferenceFontMetrics,
  type CjkLang,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet } from './types.js';
import { officeRequestKey, shapeLineFontRuns } from './shape-office-line.js';

/** Yield every textual cell value carried by the parsed workbook: the shared
 *  string table (`text` plus rich-text `runs[].text`). This is the bulk of a
 *  workbook's painted text and is present in BOTH main and worker at parse time
 *  (sheets parse lazily, but the shared string table is workbook-level).
 *  Numbers / dates carry no script-specific glyphs, so they are irrelevant to
 *  the script evidence. */
export function* xlsxTextRuns(wb: ParsedWorkbook | undefined): Generator<string> {
  for (const s of wb?.sharedStrings ?? []) {
    if (s.runs && s.runs.length > 0) {
      for (const r of s.runs) yield r.text;
    } else {
      yield s.text;
    }
  }
}

/** The same workbook hint is used for font demand and ambiguous cell fallback. */
export function xlsxCjkFallback(wb: ParsedWorkbook | undefined, fallback: CjkLang): CjkLang {
  for (const font of wb?.styles?.fonts ?? []) {
    const region = classifyCjkFont(font.name);
    if (region) return cjkFallbackForText(xlsxTextRuns(wb), region);
  }
  return cjkFallbackForText(xlsxTextRuns(wb), fallback);
}

/** Exact Calibri style slots in the workbook style and shared-string tables.
 * An omitted font name uses the workbook default Calibri chain; unrelated
 * authored families never borrow Carlito merely through a CSS fallback tail.
 * Styles are prepared before a sheet is pulled so Normal-font MDW and viewer
 * geometry cannot be captured against a different fallback resource. This can
 * prepare an unused style slot, bounded by the four supported tuples. */
export function xlsxOfficeFontRequests(wb: ParsedWorkbook | undefined): OfficeFontFallbackRequest[] {
  const found = new Map<string, OfficeFontFallbackRequest>();
  const add = (name: string | null | undefined, bold: boolean, italic: boolean) => {
    if ((name?.trim().toLowerCase() || 'calibri') !== 'calibri') return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    found.set(`${weight}:${style}`, { family: 'Calibri', weight, style });
  };
  for (const font of wb?.styles?.fonts ?? []) add(font.name, font.bold, font.italic);
  for (const shared of wb?.sharedStrings ?? []) {
    for (const run of shared.runs ?? []) {
      if (run.font) add(run.font.name, run.font.bold, run.font.italic);
    }
  }
  return [...found.values()];
}

/** Inline strings and DrawingML shapes are worksheet-local and absent from the
 * bootstrap shared-string table. Shape preflight covers every shape text run
 * whose single named face has a catalogued exact style, because each run's
 * face contributes to Excel's shape line box; cell requests remain
 * Calibri-only. The loader bounds the number of probed sources. */
export function xlsxWorksheetOfficeFontRequests(ws: Worksheet): OfficeFontFallbackRequest[] {
  const found = new Map<string, OfficeFontFallbackRequest>();
  // Column MDW uses the workbook default font, `<fonts>[0]`. Preflight that
  // authored face before measuring it, including families that occur nowhere
  // in a cell's text. Exact local bytes win over catalog references.
  const normalFamily = ws.defaultFontFamily?.trim();
  const normalWeight = ws.defaultFontBold ? 700 : 400;
  const normalStyle = ws.defaultFontItalic ? 'italic' : 'normal';
  if (normalFamily && findReferenceFontMetrics(normalFamily, { weight: normalWeight, style: normalStyle }).length) {
    const request = { family: normalFamily, weight: normalWeight, style: normalStyle } as const;
    found.set(officeRequestKey(request), request);
  }
  for (const row of ws.rows) for (const cell of row.cells) {
    if (cell.value.type !== 'text') continue;
    for (const run of cell.value.runs ?? []) {
      const font = run.font;
      if (!font || (font.name?.trim().toLowerCase() || 'calibri') !== 'calibri') continue;
      const weight = font.bold ? 700 : 400;
      const style = font.italic ? 'italic' : 'normal';
      const request = { family: 'Calibri', weight, style } as const;
      found.set(officeRequestKey(request), request);
    }
  }
  for (const anchor of ws.shapeGroups ?? []) for (const shape of anchor.shapes) {
    if (!shape.text) continue;
    for (const run of shapeLineFontRuns(shape.text)) {
      const weight = run.bold ? 700 : 400;
      const style = run.italic ? 'italic' : 'normal';
      if (findReferenceFontMetrics(run.fontFace!, { weight, style }).length === 0) continue;
      const request = { family: run.fontFace!.trim(), weight, style } as const;
      found.set(officeRequestKey(request), request);
    }
  }
  return [...found.values()];
}
