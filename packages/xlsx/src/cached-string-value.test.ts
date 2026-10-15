/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { beforeAll, expect, it } from 'vitest';
// @ts-ignore — shared synthetic ZIP fixture is a plain JavaScript module.
import { storedZip } from '../../../tests/fixtures/chart-ex-packages.mjs';
import { initSync, XlsxArchive } from './wasm/xlsx_parser.js';
import { compileCf, evaluateCf } from './conditional-format.js';
import { evaluateFormula } from './formula.js';
import type { ParsedWorkbook, Row, Worksheet } from './types.js';

beforeAll(() => {
  initSync({ module: new Uint8Array(readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url))) });
});

it('retains cached empty text through WASM, worksheet pulls and conditional formatting', () => {
  const bytes = storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      + '</Relationships>'],
    ['xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      + '</Relationships>'],
    ['xl/styles.xml', '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + '<fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>'
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
      + '<borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
      + '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">'
      + '<c r="A1" t="str"><f>""</f><v/></c><c r="B1" t="str"><f>""</f></c>'
      + '<c r="C1" s="0"/><c r="D1"><v/></c>'
      + '</row></sheetData><conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1">'
      + '<formula>COUNTA(A1)=1</formula></cfRule></conditionalFormatting></worksheet>'],
  ]) as Uint8Array;
  const archive = new XlsxArchive(bytes, undefined, undefined, undefined);
  const decode = (data: Uint8Array) => JSON.parse(new TextDecoder().decode(data));
  try {
    const workbook = decode(archive.parse()) as ParsedWorkbook;
    archive.open_sheet_cursor(0, 'Sheet1');
    const rows: Row[] = [];
    let worksheet: Worksheet | undefined;
    // One row needs only preview, rows and terminal products; bound failures.
    for (let pull = 0; pull < 4 && !worksheet; pull++) {
      const unit = decode(archive.pull_sheet_cursor(128)) as
        | { kind: 'preview' }
        | { kind: 'rows'; rows: Row[] }
        | { kind: 'finished'; worksheet: Worksheet };
      if (unit.kind === 'rows') rows.push(...unit.rows);
      if (unit.kind === 'finished') { worksheet = unit.worksheet; break; }
    }
    if (!worksheet) throw new Error('single-row worksheet did not finish');
    archive.acknowledge_sheet_cursor_terminal();
    worksheet.rows = rows;
    const cfContext = compileCf(worksheet);
    const context = { row: 1, col: 1, anchorRow: 1, anchorCol: 1,
      cellIndex: cfContext.cellIndex, definedNames: cfContext.definedNames, depth: 0 };
    const value = (formula: string) => evaluateFormula(formula, context);
    expect(value('COUNTA(A1)')).toEqual({ kind: 'value', value: 1 });
    expect(value('ISBLANK(A1)')).toEqual({ kind: 'value', value: false });
    expect(value('COUNT(A1)')).toEqual({ kind: 'value', value: 0 });
    expect(value('COUNTBLANK(A1)')).toEqual({ kind: 'value', value: 1 });
    expect(value('COUNTA(B1:D1)')).toEqual({ kind: 'value', value: 0 });
    expect(evaluateCf(cfContext.cellIndex.get('1:1'), 1, 1, cfContext, workbook.styles.dxfs).fill?.fgColor)
      .toBe('#FF0000');
  } finally {
    archive.free();
  }
});
