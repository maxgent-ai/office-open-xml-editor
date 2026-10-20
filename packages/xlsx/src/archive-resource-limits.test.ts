import { readFileSync } from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { storedZip } from '@silurus/ooxml-core/testing';
import { initSync, XlsxArchive } from './wasm/xlsx_parser.js';

function workbookWithStrings(sharedStrings: string): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
    'xl/workbook.xml': '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
    'xl/styles.xml': '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
    'xl/sharedStrings.xml': sharedStrings,
  }));
}

beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url)) });
});

describe('XlsxArchive.parse resource scope (real WASM)', () => {
  it.each([
    { kind: 'well-formed', attributes: '' },
    { kind: 'malformed', attributes: ` a='1' a='2'` },
  ])('rejects $kind over-deep shared strings and keeps the session poisoned', ({ attributes }) => {
    const archive = new XlsxArchive(workbookWithStrings(`<sst${attributes}>` + '<si>'.repeat(257) + '</si>'.repeat(257) + '</sst>'));
    try {
      let failure: unknown;
      try { archive.parse(); } catch (error) { failure = error; }
      expect(failure).toEqual(expect.stringContaining('OOXML_RESOURCE_LIMIT:'));
      const envelope = JSON.parse((failure as string).slice('OOXML_RESOURCE_LIMIT:'.length));
      expect(envelope.details.violation).toMatchObject({
        operation: 'parse', resource: 'xml-tree', metric: 'depth', limit: 256, observed: 257,
      });
      expect(() => archive.assert_healthy()).toThrow(failure as string);
      expect(() => archive.parse()).toThrow(failure as string);
    } finally {
      archive.free();
    }
  });

  it('keeps shallow malformed shared strings as a recoverable part error', () => {
    const archive = new XlsxArchive(workbookWithStrings('<sst>' + `<si a='> /' a="/>" />`.repeat(257) + '</sst>'));
    try {
      const model = JSON.parse(new TextDecoder().decode(archive.parse()));
      expect(model.sharedStrings).toEqual([]);
      expect(model.workbook.parseError).toContain('xl/sharedStrings.xml');
      expect(() => archive.assert_healthy()).not.toThrow();
    } finally {
      archive.free();
    }
  });
});
