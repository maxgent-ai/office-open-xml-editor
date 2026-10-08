/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { createLayoutServices } from '../layout-runtime.js';
import { layoutDocument } from '../document-layout.js';
import type { TableLayout } from './types.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

function layoutTable(tableXml: string): TableLayout {
  const parts = new Map([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    </Types>`],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`],
    ['word/_rels/document.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    </Relationships>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body>${tableXml}
      <w:sectPr><w:pgSz w:w="9360" w:h="14400"/>
        <w:pgMar w:left="360" w:right="360" w:top="1440" w:bottom="1440"/>
      </w:sectPr></w:body></w:document>`],
    ['word/settings.xml', `<w:settings xmlns:w="${W}"><w:compat>
      <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
    </w:compat></w:settings>`],
  ]);
  const archive = new DocxArchive(storeZip(new Map(
    [...parts].map(([name, xml]) => [name, encoder.encode(xml)]),
  )));
  try {
    const model = normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
    const context = {
      font: '10px serif', letterSpacing: '0px', fontKerning: 'none',
      measureText: (text: string) => ({
        width: [...text].length * 6,
        fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      }),
    } as unknown as CanvasRenderingContext2D;
    const result = layoutDocument(model, createLayoutServices(model, { measureContext: context }), {
      currentDateMs: 0,
    });
    const table = result.pages[0]?.layers.body.find((block) => block.kind === 'table');
    if (table?.kind !== 'table') throw new Error('expected retained parsed table');
    return table;
  } finally {
    archive.free();
  }
}

const cell = '<w:tc><w:tcPr><w:tcW w:type="auto" w:w="0"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>';

function widthTable(bodyWidth: string, exceptionWidth?: string): string {
  // Both maxima exceed the band, making the occurrence ceiling observable.
  const growingCell = cell.replace('>cell<', `>${'cell '.repeat(60).trim()}<`);
  return `<w:tbl><w:tblPr>${bodyWidth}<w:tblInd w:type="dxa" w:w="108"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="4400"/><w:gridCol w:w="4400"/></w:tblGrid>
    <w:tr>${exceptionWidth ? `<w:tblPrEx>${exceptionWidth}</w:tblPrEx>` : ''}${growingCell}${growingCell}</w:tr>
  </w:tbl>`;
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('parsed AutoFit occurrence width constraints', () => {
  const auto = '<w:tblW w:type="auto" w:w="0"/>';
  const preferred = '<w:tblW w:type="dxa" w:w="8800"/>';

  it('first-row auto clears the body dxa width for both the ceiling and solver', () => {
    const overridden = layoutTable(widthTable(preferred, auto));
    const direct = layoutTable(widthTable(auto));
    expect(overridden.flowBounds.widthPt).toBeCloseTo(426.6, 8);
    expect(overridden.columnWidthsPt).toEqual(direct.columnWidthsPt);
  });

  it('first-row dxa replaces the body auto width for both the ceiling and solver', () => {
    const overridden = layoutTable(widthTable(auto, preferred));
    const direct = layoutTable(widthTable(preferred));
    expect(overridden.flowBounds.widthPt).toBeCloseTo(440, 8);
    expect(overridden.columnWidthsPt).toEqual(direct.columnWidthsPt);
  });

  it.each(['Before', 'After'] as const)(
    'retains grid%s and its specified row width when placing the parsed cell',
    (edge) => {
      // The omitted track starts at 25pt in tblGrid. Its 50pt row preference
      // must survive; it has no content maximum and cannot be a growth track.
      const grid = edge === 'Before' ? [500, 1000] : [1000, 500];
      const table = layoutTable(`<w:tbl><w:tblPr>${auto}</w:tblPr>
        <w:tblGrid>${grid.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
        <w:tr><w:trPr><w:grid${edge} w:val="1"/><w:w${edge} w:type="dxa" w:w="1000"/></w:trPr>
          ${cell}</w:tr></w:tbl>`);
      expect(table.columnWidthsPt).toEqual([50, 50]);
      expect(table.flowBounds.widthPt).toBe(100);
      expect(table.rows[0]?.cells[0]?.flowBounds.xPt).toBe(edge === 'Before' ? 68 : 18);

      // The row preference still excludes this track from content growth
      // when another row has an unpreferred cell in the same shared column.
      const mixed = layoutTable(`<w:tbl><w:tblPr>${auto}</w:tblPr>
        <w:tblGrid>${grid.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>
        <w:tr>${cell}${cell}</w:tr>
        <w:tr><w:trPr><w:grid${edge} w:val="1"/><w:w${edge} w:type="dxa" w:w="1000"/></w:trPr>
          ${cell}</w:tr></w:tbl>`);
      expect(mixed.columnWidthsPt).toEqual([50, 50]);
      expect(mixed.rows[1]?.cells[0]?.flowBounds.xPt).toBe(edge === 'Before' ? 68 : 18);
    },
  );

  it('preserves a noWrap dxa preference through the initial tblW fit', () => {
    const preferredCell = (noWrap: boolean, text = 'x') => cell
      .replace('w:type="auto" w:w="0"', 'w:type="dxa" w:w="1600"')
      .replace('</w:tcPr>', `${noWrap ? '<w:noWrap/>' : ''}</w:tcPr>`)
      .replace('>cell<', `>${text}<`);
    const table = (first: string, second: string) => `<w:tbl><w:tblPr>
      <w:tblW w:type="dxa" w:w="2000"/>
      <w:tblCellMar><w:left w:type="dxa" w:w="0"/><w:right w:type="dxa" w:w="0"/></w:tblCellMar></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1600"/><w:gridCol w:w="1600"/></w:tblGrid>
      <w:tr>${first}${second}</w:tr></w:tbl>`;
    expect(layoutTable(table(preferredCell(true), preferredCell(false))).columnWidthsPt)
      .toEqual([80, 20]);
    expect(layoutTable(table(preferredCell(false), preferredCell(false))).columnWidthsPt)
      .toEqual([50, 50]);
    const fixed = table(preferredCell(true), preferredCell(false))
      .replace('</w:tblPr>', '<w:tblLayout w:type="fixed"/></w:tblPr>');
    expect(layoutTable(fixed).columnWidthsPt).toEqual([50, 50]);
    // Once the neighbour reaches its absolute minimum (7 * 6pt glyph
    // advances, with explicit zero margins), the preference can shrink.
    const constrained = layoutTable(table(preferredCell(true), preferredCell(false, 'xxxxxxx')));
    expect(constrained.columnWidthsPt[0]).toBeCloseTo(58, 8);
    expect(constrained.columnWidthsPt[1]).toBeCloseTo(42, 8);
  });

  it('preserves a spanning noWrap dxa preference during the initial tblW fit', () => {
    const spanning = cell.replace('w:type="auto" w:w="0"', 'w:type="dxa" w:w="1600"')
      .replace('</w:tcPr>', '<w:gridSpan w:val="2"/><w:noWrap/></w:tcPr>');
    const neighbour = cell.replace('w:type="auto" w:w="0"', 'w:type="dxa" w:w="1600"')
      .replace('>cell<', '>x<');
    const table = layoutTable(`<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="2000"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="400"/><w:gridCol w:w="1600"/></w:tblGrid>
      <w:tr>${spanning}${neighbour}</w:tr></w:tbl>`);
    expect(table.columnWidthsPt).toEqual([60, 20, 20]);
  });

  it.each(['dxa', 'nil', 'pct'] as const)('treats noWrap tcW %s zero as auto in measurement and fitting', (kind) => {
    const noWrapCell = cell.replace('</w:tcPr>', '<w:noWrap/></w:tcPr>')
      .replace('>cell<', `>${'alpha beta gamma delta '.repeat(20).trim()}<`);
    const table = (first: string) => `<w:tbl><w:tblPr>${auto}
      <w:tblInd w:type="dxa" w:w="108"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4532"/></w:tblGrid>
      <w:tr>${first}${cell}</w:tr></w:tbl>`;
    const automatic = layoutTable(table(noWrapCell));
    const equivalent = layoutTable(table(noWrapCell.replace('w:type="auto"', `w:type="${kind}"`)));
    expect(automatic.columnWidthsPt[0]).toBeGreaterThan(350);
    expect(equivalent.columnWidthsPt).toEqual(automatic.columnWidthsPt);
    const lines = (layout: TableLayout) => layout.rows[0]!.cells.map((c) => c.blocks.flatMap((b) =>
      b.layout.kind === 'paragraph' ? b.layout.lines.map((l) => l.placements
        .filter((placement) => placement.kind === 'text').map((placement) => placement.text).join('')) : []));
    expect(lines(equivalent)).toEqual(lines(automatic));
    if (kind === 'nil') {
      // nil ignores the numeric width even when it is nonzero.
      const nilWidth = noWrapCell.replace('w:type="auto" w:w="0"', 'w:type="nil" w:w="1600"');
      const nonzeroNil = layoutTable(table(nilWidth));
      expect(nonzeroNil.columnWidthsPt).toEqual(automatic.columnWidthsPt);
      expect(lines(nonzeroNil)).toEqual(lines(automatic));
    }
  });

  it('retains a saved track with no cell even without an explicit row skip element', () => {
    const table = layoutTable(`<w:tbl><w:tblPr>${auto}</w:tblPr>
      <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
      <w:tr>${cell}</w:tr></w:tbl>`);
    expect(table.columnWidthsPt).toEqual([50, 50]);
    expect(table.flowBounds.widthPt).toBe(100);
  });
});
