import { afterEach, expect, it, vi } from 'vitest';
import { OoxmlError, type ModelSource } from '@silurus/ooxml-core';
import { openModelSourceModule } from '@silurus/ooxml-core/internal/model-source';
import { MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml/internal/companion';
import { serializeWorkerError } from '@silurus/ooxml-core/worker';
import { buildCfbFixture } from '@silurus/ooxml-core/testing';
import { DocxDocument } from '../../docx/src/document.js';
import { XlsxWorkbook } from '../../xlsx/src/workbook.js';
import { PptxPresentation } from '../../pptx/src/presentation.js';
import { openDocxDocument, openXlsxWorkbook, openPptxPresentation, skiaFactory } from './node/node-facade.js';
import { testDocSource, testXlsSource, testPptSource, TEST_SOURCE_URLS } from './test-sources.js';

const cases = [
  { format: 'doc', target: 'docx', stream: 'WordDocument', source: testDocSource },
  { format: 'xls', target: 'xlsx', stream: 'Workbook', source: testXlsSource },
  { format: 'ppt', target: 'pptx', stream: 'PowerPoint Document', source: testPptSource },
] as const;

afterEach(() => vi.unstubAllGlobals());

// Real native readers and source modules in an in-memory Worker transport.
// Both browser modes still parse in a Worker. This pins each public facade's
// wire decoder and failure cleanup without starting a browser for unit tests.
it.each(cases.flatMap(test => ['main', 'worker'].map(mode => ({ ...test, mode }))))(
  '$format rejection reaches browser $mode callers as a viewer-owned OoxmlError', async ({ format, target, stream, source, mode }) => {
    let terminated = 0;
    class ParserWorker {
      listeners = new Set<(event: MessageEvent) => void>();
      addEventListener(type: string, listener: (event: MessageEvent) => void) { if (type === 'message') this.listeners.add(listener); }
      removeEventListener(_type: string, listener: (event: MessageEvent) => void) { this.listeners.delete(listener); }
      terminate() { terminated++; }
      postMessage(message: { type?: string; kind?: string; id: number }) {
        if ((message.type ?? message.kind) !== 'parse') return;
        queueMicrotask(async () => {
          try {
            await openModelSourceModule({ protocol: MODEL_SOURCE_MODULE_PROTOCOL, target,
              moduleUrl: TEST_SOURCE_URLS[format].moduleUrl, config: { wasmUrl: TEST_SOURCE_URLS[format].wasmUrl } },
            new Uint8Array(buildCfbFixture(['Root Entry', stream])), archive => archive);
            throw new Error('expected malformed native input to reject');
          } catch (error) {
            const payload = structuredClone({ ...(message.kind ? { kind: 'error' } : { type: 'error' }), id: message.id,
              ...serializeWorkerError(error) });
            for (const listener of this.listeners) listener({ data: payload } as MessageEvent);
          }
        });
      }
    }
    vi.stubGlobal('Worker', ParserWorker);
    vi.stubGlobal('OffscreenCanvas', class {});
    vi.stubGlobal('location', { href: 'https://example.test/' });
    const input = buildCfbFixture(['Root Entry', stream]);
    const options = { mode: mode as 'main' | 'worker' };
    const pending = target === 'docx' ? DocxDocument.load(input, { ...options, legacyDoc: source() as ModelSource<'docx'> })
      : target === 'xlsx' ? XlsxWorkbook.load(input, { ...options, legacyXls: source() as ModelSource<'xlsx'> })
        : PptxPresentation.load(input, { ...options, legacyPpt: source() as ModelSource<'pptx'> });
    await expect(pending).rejects.toBeInstanceOf(OoxmlError);
    await expect(pending).rejects.toMatchObject({ code: 'unsupported-content', details: { format, reason: 'source-subset' } });
    expect(terminated).toBeGreaterThan(0);
  },
);

it.each(cases)('$format rejection reaches the Node opener as the same typed contract', async ({ format, target, stream, source }) => {
  const input = buildCfbFixture(['Root Entry', stream]);
  const pending = target === 'docx'
    ? openDocxDocument(input, { legacyDoc: source() as ModelSource<'docx'>, factory: skiaFactory() })
    : target === 'xlsx' ? openXlsxWorkbook(input, { legacyXls: source() as ModelSource<'xlsx'> })
      : openPptxPresentation(input, { legacyPpt: source() as ModelSource<'pptx'> });
  await expect(pending).rejects.toBeInstanceOf(OoxmlError);
  await expect(pending).rejects.toMatchObject({ code: 'unsupported-content', details: { format, reason: 'source-subset' } });
});
