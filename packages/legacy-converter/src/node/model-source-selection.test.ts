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

  // MS-DOC 2.6.4 sprmSTextFlow (0x5033) carries an MS-ODRAW 2.4.5 MSOTXFL.
  // Every valid value is admitted. The public direction is a display family:
  // 0 and 4 horizontal, 1 the established tbRl, 3 and 5 the all-rotated
  // clockwise btLr family (a display equivalence, not a Word export claim).
  // The effective raw value is retained in the private placement wire, absent
  // when the section has no sprmSTextFlow.
  const textFlow = (value: number) => concat(little16(0x5033), little16(value));
  type SectionFacts = { type?: string; textDirection?: string | null; __sectionPlacement?: Record<string, unknown> };
  const loadDoc = (bytes: Uint8Array) => materializeDocxDocument(bytes, { modelSources: [testDocSource()] });
  const finalSection = (document: { section: unknown }) => document.section as SectionFacts;
  const endingSection = (document: { body: unknown[] }) => {
    const breaks = (document.body as SectionFacts[]).filter(element => element.type === 'sectionBreak');
    expect(breaks).toHaveLength(1);
    return breaks[0]!;
  };
  const direction = (section: SectionFacts) => section.textDirection ?? null;
  const rawFlow = (section: SectionFacts) => {
    expect(section.__sectionPlacement).toBeTypeOf('object');
    return section.__sectionPlacement!.nativeTextFlow;
  };
  const oneSection = (properties: Uint8Array) => buildDocFixture({ text: 'Body\r', sectionProperties: properties });
  const twoSections = (first: Uint8Array, last: Uint8Array) =>
    buildDocFixture({ text: 'One\fTwo\r', sectionEnds: [4, 8], sectionProperties: [first, last] });
  const none = new Uint8Array();

  it.each([
    { flow: 0, expected: null },
    { flow: 1, expected: 'tbRl' },
    { flow: 3, expected: 'btLr' },
    { flow: 4, expected: null },
    { flow: 5, expected: 'btLr' },
  ])('project native section text flow $flow to its display family in final and ending sections', async ({ flow, expected }) => {
    expect(direction(finalSection(await loadDoc(oneSection(textFlow(flow)))))).toBe(expected);
    const sectioned = await loadDoc(twoSections(textFlow(flow), none));
    expect(direction(endingSection(sectioned))).toBe(expected);
    expect(direction(finalSection(sectioned))).toBeNull();
  });

  it('retain the effective raw section text flow, distinguishing absence from an explicit 0', async () => {
    expect(rawFlow(finalSection(await loadDoc(oneSection(none))))).toBeUndefined();
    for (const flow of [0, 1, 2, 3, 4, 5]) {
      expect(rawFlow(finalSection(await loadDoc(oneSection(textFlow(flow))))), `MSOTXFL ${flow}`).toBe(flow);
    }
    const sectioned = await loadDoc(twoSections(textFlow(2), none));
    expect(rawFlow(endingSection(sectioned))).toBe(2);
    expect(rawFlow(finalSection(sectioned))).toBeUndefined();
  });

  it('apply the last section text flow modifier', async () => {
    for (const { modifiers, expected } of [
      { modifiers: [1, 3, 0], expected: null },
      { modifiers: [0, 3], expected: 'btLr' },
      { modifiers: [3, 1], expected: 'tbRl' },
    ]) {
      const section = finalSection(await loadDoc(oneSection(concat(...modifiers.map(textFlow)))));
      expect(direction(section), modifiers.join('->')).toBe(expected);
      expect(rawFlow(section), modifiers.join('->')).toBe(modifiers.at(-1));
    }
  });

  it('keep each section text flow with its own section', async () => {
    let sectioned = await loadDoc(twoSections(textFlow(2), textFlow(1)));
    expect([rawFlow(endingSection(sectioned)), rawFlow(finalSection(sectioned))]).toEqual([2, 1]);
    expect(direction(finalSection(sectioned))).toBe('tbRl');
    sectioned = await loadDoc(twoSections(textFlow(3), textFlow(0)));
    expect([direction(endingSection(sectioned)), direction(finalSection(sectioned))]).toEqual(['btLr', null]);
    sectioned = await loadDoc(twoSections(none, textFlow(5)));
    expect([direction(endingSection(sectioned)), direction(finalSection(sectioned))]).toEqual([null, 'btLr']);
    expect(rawFlow(endingSection(sectioned))).toBeUndefined();
  });

  it('keep a value outside MSOTXFL invalid', async () => {
    await expect(loadDoc(oneSection(textFlow(6))))
      .rejects.toThrow(new Error('UNSUPPORTED:invalid Word section text flow'));
  });

  it('keep horizontal-in-vertical text admitted only in tbRl sections', async () => {
    // MS-DOC 2.6.1 sprmCFELayout: FarEastLayoutOperand cb 6, UFEL fTNY, ID.
    const tny = concat(little16(0xca78), new Uint8Array([6, 1, 0, 0, 0, 0, 0]));
    const withFlow = (flow: number) => buildDocFixture({ text: 'Body\r', characterProperties: tny, sectionProperties: textFlow(flow) });
    const vertical = await loadDoc(withFlow(1));
    expect(JSON.stringify(vertical.body)).toContain('"eastAsianVert":true');
    for (const flow of [0, 2, 3, 4, 5]) {
      await expect(loadDoc(withFlow(flow)), `MSOTXFL ${flow}`).rejects.toThrow(unsupportedFormatting);
    }
  });

  it('reject an oversize claimed input before opening it', async () => {
    const bytes = buildDocFixture();
    await expect(materializeDocxDocument(bytes, { modelSources: [testDocSource({ maxInputBytes: bytes.byteLength - 1 })] }))
      .rejects.toThrow(RangeError);
  });
});
