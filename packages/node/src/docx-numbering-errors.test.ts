import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { storedZip } from '@silurus/ooxml-core/testing';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cases = [
  { reason: 'whitespace', value: ' 8 ', definition: 0,
    message: 'Word cannot open this DOCX numbering level (whitespace).' },
  { reason: 'non-decimal', value: 'abc', definition: 0,
    message: 'Word cannot open this DOCX numbering level (non-decimal).' },
  { reason: 'level-definition', value: '0', definition: 9,
    message: 'The DOCX defines a numbering level outside Word’s supported 0–8 range and requires repair.' },
];
function fixture(value: string, definition: number) {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'],
    ['word/document.xml', `<w:document xmlns:w="${w}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="${value}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p></w:body></w:document>`],
    ['word/numbering.xml', `<w:numbering xmlns:w="${w}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="${definition}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`],
  ]);
}
const workerEntries = {
  main: () => import('../../docx/src/worker.ts'),
  worker: () => import('../../docx/src/render-worker.ts'),
  source: () => import('../../docx/src/worker-source.ts'),
};
afterEach(() => { vi.unstubAllGlobals(); });

// Main mode uses the inline parser worker; worker mode uses the render worker.
// Load the real entries and deserialize their real wire payloads, replacing only
// the browser message surface. No replica of their parser/catch logic is used.
type WorkerScope = {
  onmessage: null | ((event: MessageEvent) => Promise<void>);
  postMessage(message: Record<string, unknown>): void;
};
const scopes = new Map<string, WorkerScope>();
beforeAll(async () => {
  // Keep module compilation outside the request assertions: the render worker
  // imports the complete layout/paint graph, even though parsing rejects first.
  for (const [mode, entry] of Object.entries(workerEntries)) {
    const scope: WorkerScope = { onmessage: null, postMessage() {} };
    vi.stubGlobal('self', scope);
    await entry();
    scopes.set(mode, scope);
  }
}, 120_000);

it.each(Object.keys(workerEntries))('keeps every numbering code and message in the %s entry', async mode => {
  const messages: Record<string, unknown>[] = [];
  const scope = scopes.get(mode)!;
  scope.postMessage = message => { messages.push(message); };
  vi.stubGlobal('self', scope);
  const dispatch = async (data: unknown) => { await scope.onmessage!({ data } as MessageEvent); };
  const wasm = readFileSync(new URL('../../docx/src/wasm/docx_parser_bg.wasm', import.meta.url));
  await dispatch({ type: 'init', wasmUrl: `data:application/wasm;base64,${wasm.toString('base64')}` });
  const { deserializeWorkerError, normalizeResourcePolicy } = await import('@silurus/ooxml-core/worker');
  const { OoxmlError } = await import('@silurus/ooxml-core');
  for (const { value, definition, message } of cases) {
    messages.length = 0;
    await dispatch({ type: 'parse', id: 1, data: fixture(value, definition).buffer,
      defaultCurrentDateMs: 0, resourcePolicy: normalizeResourcePolicy({}) });
    expect(messages).toHaveLength(1);
    expect(messages[0]).toMatchObject({ type: 'error', id: 1 });
    const error = deserializeWorkerError(messages[0] as never);
    expect(error).toBeInstanceOf(OoxmlError);
    expect(error).toMatchObject({ code: 'invalid-numbering', message });
  }
});

it('keeps every numbering code and message in the Node materializer', async () => {
  const { materializeDocxDocument } = await import('./docx.ts');
  const { OoxmlError } = await import('@silurus/ooxml-core');
  for (const { value, definition, message } of cases) {
    const loading = materializeDocxDocument(fixture(value, definition));
    await expect(loading).rejects.toBeInstanceOf(OoxmlError);
    await expect(loading).rejects.toMatchObject({ code: 'invalid-numbering', message });
  }
}, 20_000);

it('normalizes numbering rejection while a selected DOCX model source opens', async () => {
  const { materializeDocxDocument } = await import('./docx.ts');
  const sample = cases[0];
  const loading = materializeDocxDocument(fixture(sample.value, sample.definition), {
    legacyDoc: {
      target: 'docx', claim: () => true,
      beginLoad: () => ({
        module: {
          protocol: 'ooxml-model-source-module/v2', target: 'docx',
          moduleUrl: new URL('./test-fixtures/docx-model-source.mjs', import.meta.url).href,
          config: { openError: 'OOXML_DOCX_ILVL:cannot-open:whitespace' },
        },
        release() {},
      }),
    },
  });
  await expect(loading).rejects.toMatchObject({ code: 'invalid-numbering', message: sample.message });
});
