import type { DocParagraph, DocxDocumentModel } from '../types.js';
import { syntheticDocxModel } from './synthetic-document.js';

/** Plain shared-model wire fixture; native control/guard decoding is exercised
 * by the native producer's tests, independently of this layout contract. */
export function noteContinuationModel(): DocxDocumentModel {
  const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
  model.section = { ...model.section, pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginBottom: 10, marginLeft: 10, marginRight: 10 };
  const paragraph = model.body[0] as DocParagraph;
  const run = paragraph.runs[0];
  if (run?.type !== 'text') throw new Error('Expected text fixture');
  model.footnotes = [{ id: '7', content: Array.from({ length: 30 }, (_, index) => ({
    ...paragraph, type: 'paragraph', runs: [{ ...run, text: `note-${index};` }],
  })) }];
  paragraph.runs = [{ ...run, text: '1', vertAlign: 'super', noteRef: { kind: 'footnote', id: '7' } }];
  return JSON.parse(JSON.stringify({ ...model, __noteLayoutSettings: {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'short',
  } })) as DocxDocumentModel;
}

/** Exact 10pt lines on an 80pt body region (no separator story, so the 6pt
 * band): five filler paragraphs, then `B01` opened by reference 1 with text
 * after it; note 1 is four one-line paragraphs. Continued: 50 + 10 + band 6 +
 * one note line = 76 keeps `note-0;` beside B01 and carries the other three.
 * Whole notes: 60 + 46 > 80, so B01 and the whole note move to page 2
 * (10 + 46 <= 80). */
export function noteContinuationBoundaryModel(): DocxDocumentModel {
  const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
  model.section = { ...model.section, pageWidth: 200, pageHeight: 100,
    marginTop: 10, marginBottom: 10, marginLeft: 10, marginRight: 10 };
  const base = model.body[0] as DocParagraph;
  const run = base.runs[0];
  if (run?.type !== 'text') throw new Error('Expected text fixture');
  const exact = (...texts: string[]) => ({
    ...base, type: 'paragraph', spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: 10, rule: 'exact', explicit: true },
    runs: texts.map(text => ({ ...run, text })),
  }) as DocParagraph;
  const reference = exact('1', 'B01');
  reference.runs = [{ ...run, text: '1', vertAlign: 'super', noteRef: { kind: 'footnote', id: '7' } },
    reference.runs[1]!];
  model.body = [...Array.from({ length: 5 }, (_, index) => exact(`F${index}`)), reference] as unknown as DocxDocumentModel['body'];
  model.footnotes = [{ id: '7',
    content: Array.from({ length: 4 }, (_, index) => exact(`note-${index};`)) as unknown as DocxDocumentModel['body'] }];
  return JSON.parse(JSON.stringify({ ...model, __noteLayoutSettings: {
    footnoteSeparator: 'short', footnoteContinuationSeparator: 'short',
  } })) as DocxDocumentModel;
}
