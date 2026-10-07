import { describe, expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { buildDocFixture, buildPptFixture, buildXlsFixture, concat, little16 } from '../test-fixtures.js';
import { testDocSource, testPptSource, testXlsSource } from '../test-sources.js';
import {
  materializeDocxDocument,
  materializePptxPresentation,
  materializeXlsxWorkbook,
  openPptxPresentation,
  openXlsxWorkbook,
} from './node-facade.js';

const cfb = (...streams: string[]) => new Uint8Array(buildCfbFixture(['Root Entry', ...streams]));
const legacyRejection = { code: 'legacy-binary-format' };

describe('Node openers and legacy model sources', () => {
  it('keeps the original Node OOXML rejection without model sources', async () => {
    for (const open of [
      () => materializeDocxDocument(cfb('WordDocument')),
      () => openXlsxWorkbook(cfb('Workbook')),
      () => openPptxPresentation(cfb('PowerPoint Document')),
    ]) {
      await expect(open()).rejects.toBeInstanceOf(OoxmlError);
      await expect(open()).rejects.toMatchObject({ code: 'not-ooxml' });
    }
  });

  it('leave foreign-family and ambiguous containers to the OOXML rejection', async () => {
    const doc = { modelSources: [testDocSource()] };
    const xls = { modelSources: [testXlsSource()] };
    const ppt = { modelSources: [testPptSource()] };
    await expect(materializeDocxDocument(cfb('Workbook'), doc)).rejects.toMatchObject(legacyRejection);
    await expect(materializeDocxDocument(cfb('WordDocument', 'Workbook'), doc)).rejects.toMatchObject(legacyRejection);
    await expect(openXlsxWorkbook(cfb('PowerPoint Document'), xls)).rejects.toMatchObject(legacyRejection);
    await expect(openPptxPresentation(cfb('WordDocument'), ppt)).rejects.toMatchObject(legacyRejection);
    // An encrypted package is never claimed; the OOXML path reports it.
    await expect(materializeDocxDocument(cfb('EncryptionInfo', 'EncryptedPackage', 'WordDocument'), doc))
      .rejects.toMatchObject({ code: 'encrypted' });
    // A source for another target is a configuration error, not a fallback.
    await expect(materializeDocxDocument(cfb('WordDocument'), xls as never)).rejects.toThrow(TypeError);
  });

  it('route each own family to its direct reader', async () => {
    const document = await materializeDocxDocument(buildDocFixture(), { modelSources: [testDocSource()] });
    expect(JSON.stringify(document.body)).toContain('Hello 日本語');

    const workbook = await materializeXlsxWorkbook(buildXlsFixture(), { modelSources: [testXlsSource()] });
    expect(JSON.stringify(workbook.workbookIndex.workbook)).toContain('表計算');
    expect(workbook.worksheets).toHaveLength(1);
    const cells = JSON.stringify([workbook.workbookIndex.sharedStrings, workbook.worksheets[0]]);
    expect(cells).toContain('42.5');
    expect(cells).toContain('日本語');

    // The direct PPT reader does not project a slide with outline text but no
    // drawing; reaching its fail-closed rejection proves the routing.
    await expect(materializePptxPresentation(buildPptFixture(), { modelSources: [testPptSource()] }))
      .rejects.toThrow(/UNSUPPORTED:.*no drawing/);
  });

  // MS-DOC 2.6.1 sprmCFRMarkDel (0x0800, operand 1) is deleted-revision
  // character formatting that the direct reader does not support: its atomic
  // gate rejects the whole model. A break control is a character with its own
  // piece formatting, so formatting confined to that one control must reach the
  // same gate as text, a tab or a line break, not be emitted as a plain break.
  const deletedRevision = concat(little16(0x0800), new Uint8Array([1]));
  const confinedToControl = (text: string) => buildDocFixture({ text, formattingRuns: [
    { end: 1, properties: new Uint8Array() },
    { end: 2, properties: deletedRevision },
    { end: 4, properties: new Uint8Array() },
  ] });
  const unsupportedFormatting = new Error('UNSUPPORTED:direct DOC model encountered unsupported formatting');
  const bodyTypes = (document: { body: unknown[] }) => (document.body as { type: string }[]).map(element => element.type);

  it('reject unsupported character formatting confined to a line-break control', async () => {
    await expect(materializeDocxDocument(confinedToControl('A\u000bB\r'), { modelSources: [testDocSource()] }))
      .rejects.toThrow(unsupportedFormatting);
  });

  it.each([
    { kind: 'page', control: '\f' },
    { kind: 'column', control: '\u000e' },
  ])('reject unsupported character formatting confined to a $kind break', async ({ kind, control }) => {
    const source = testDocSource();
    await expect(materializeDocxDocument(confinedToControl(`A${control}B\r`), { modelSources: [source] }))
      .rejects.toThrow(unsupportedFormatting);
    // The gate rejects this input only: the same source still opens a valid
    // document whose plain break of the same kind is kept in the model.
    const plain = await materializeDocxDocument(buildDocFixture({ text: `A${control}B\r` }), { modelSources: [source] });
    expect(bodyTypes(plain)).toContain(`${kind}Break`);
  });

  it('reject an oversize claimed input before opening it', async () => {
    const bytes = buildDocFixture();
    await expect(materializeDocxDocument(bytes, { modelSources: [testDocSource({ maxInputBytes: bytes.byteLength - 1 })] }))
      .rejects.toThrow(RangeError);
  });
});
