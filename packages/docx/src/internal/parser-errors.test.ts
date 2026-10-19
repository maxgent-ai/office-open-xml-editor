import { expect, it } from 'vitest';
import { OoxmlError } from '@silurus/ooxml-core';
import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { withDocxParserErrors } from './parser-errors.js';

it.each([
  ['cannot-open:whitespace', 'Word cannot open this DOCX numbering level (whitespace).'],
  ['cannot-open:non-decimal', 'Word cannot open this DOCX numbering level (non-decimal).'],
  ['repair-required:level-definition', 'The DOCX defines a numbering level outside Word’s supported 0–8 range and requires repair.'],
])('preserves the exact %s rejection for string and Error WASM failures', (reason, message) => {
  for (const input of [`OOXML_DOCX_ILVL:${reason}`, new Error(`OOXML_DOCX_ILVL:${reason}`)]) {
    let normalized: unknown;
    try { withDocxParserErrors(() => { throw input; }); } catch (error) { normalized = error; }
    expect(normalized).toBeInstanceOf(OoxmlError);
    expect(normalized).toMatchObject({ code: 'invalid-numbering', message });
    const restored = deserializeWorkerError(structuredClone(serializeWorkerError(normalized)));
    expect(restored).toBeInstanceOf(OoxmlError);
    expect(restored).toMatchObject({ code: 'invalid-numbering', message });
  }
});

it('preserves unrelated failures and unknown or wrapped numbering reasons by identity', () => {
  const unreadable = { toString() { throw new Error('unreadable'); } };
  const typed = new OoxmlError('invalid-numbering', 'Already normalized.');
  for (const input of [unreadable, typed, new Error('unrelated'), 'OOXML_DOCX_ILVL:unexpected',
    new Error('OOXML_DOCX_ILVL:cannot-open:empty'),
    new Error('wrapped OOXML_DOCX_ILVL:cannot-open:whitespace')]) {
    let failure: unknown;
    try { withDocxParserErrors(() => { throw input; }); } catch (error) { failure = error; }
    expect(failure).toBe(input);
  }
});
