import { describe, expect, it } from 'vitest';
import { DocxDocument } from './document.js';
import { recordResourceDiagnostics } from './resource-diagnostics.js';

const parse = { code: 'UNSUPPORTED_FEATURE', severity: 'warning', message: 'Unsupported VML',
  source: { story: 'body', storyInstance: 'body', path: [0] } } as const;
const font = { code: 'MISSING_RESOURCE', severity: 'warning', message: 'Embedded font resource limit' } as const;
const picture = { code: 'MISSING_RESOURCE', severity: 'warning', message: 'Chart picture unavailable' } as const;

describe('document diagnostics', () => {
  it.each(['main', 'worker'])('%s exposes parser, font and paint diagnostics without mutable publication state', mode => {
    const doc = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(doc, { _mode: mode, _meta: mode === 'worker'
      ? { diagnostics: [parse], resourceDiagnostics: [font] } : null,
    _getLayout: () => ({ pages: [], diagnostics: [parse] }) });
    if (mode === 'main') recordResourceDiagnostics(doc, [font]);
    recordResourceDiagnostics(doc, [picture]);
    recordResourceDiagnostics(doc, [picture]);
    expect(doc.diagnostics).toEqual([parse, font, picture]);
    expect(Object.isFrozen(doc.diagnostics[0]?.source?.path)).toBe(true);
  });
});
