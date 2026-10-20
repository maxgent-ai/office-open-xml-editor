import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { storedZip } from '@silurus/ooxml-core/testing';

const relationship = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const r = `xmlns:r="${relationship}"`;
const rels = (entries: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="${relationship}/${type}" Target="${target}"/>`;
type Format = 'docx' | 'pptx' | 'xlsx';

// Literal paths keep every dependency visible to the boundary check. Defer
// evaluation until after resetModules and browser stubs install the worker realm.
const workerEntries = {
  docx: {
    'render-worker': () => import('../../docx/src/render-worker.ts'),
    worker: () => import('../../docx/src/worker.ts'),
    'worker-source': () => import('../../docx/src/worker-source.ts'),
  },
  pptx: {
    'render-worker': () => import('../../pptx/src/render-worker.ts'),
    worker: () => import('../../pptx/src/worker.ts'),
    'worker-source': () => import('../../pptx/src/worker-source.ts'),
  },
  xlsx: {
    'render-worker': () => import('../../xlsx/src/render-worker.ts'),
    worker: () => import('../../xlsx/src/worker.ts'),
    'worker-source': () => import('../../xlsx/src/worker-source.ts'),
  },
};
const wasmBindings = {
  docx: () => import('../../docx/src/wasm/docx_parser.js'),
  pptx: () => import('../../pptx/src/wasm/pptx_parser.js'),
  xlsx: () => import('../../xlsx/src/wasm/xlsx_parser.js'),
};

// Synthetic, redistributable packages. The image is deliberately undecodable:
// extraction must fail first, before a renderer can treat decoding as optional.
function fixture(format: Format, fonts = false, empty = false): Uint8Array<ArrayBuffer> {
  const files: Array<[string, string]> = [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="fntdata" ContentType="application/x-font-ttf"/></Types>'],
  ];
  const fontRefs = [1, 2, 3];
  const picture = `<a:blipFill><a:blip r:embed="img"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
  if (format === 'pptx') {
    files.push(
      ['ppt/presentation.xml', `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ${r}><p:sldIdLst>${empty ? '' : '<p:sldId id="256" r:id="slide"/>'}</p:sldIdLst><p:sldSz cx="914400" cy="914400"/>${fonts ? `<p:embeddedFontLst>${fontRefs.map(i => `<p:embeddedFont><p:font typeface="Font ${i}"/><p:regular r:id="f${i}"/></p:embeddedFont>`).join('')}</p:embeddedFontLst>` : ''}</p:presentation>`],
      ['ppt/_rels/presentation.xml.rels', rels(rel('slide', 'slide', 'slides/slide1.xml') + (fonts ? fontRefs.map(i => rel(`f${i}`, 'font', `fonts/font${i}.fntdata`)).join('') : ''))],
      ['ppt/slides/slide1.xml', `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${a}" ${r}><p:cSld><p:spTree>${fonts ? '' : `<p:pic><p:nvPicPr><p:cNvPr id="1" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>${picture.replaceAll('a:blipFill', 'p:blipFill')}<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr></p:pic>`}</p:spTree></p:cSld></p:sld>`],
      ['ppt/slides/_rels/slide1.xml.rels', rels(rel('img', 'image', '../media/limit.png'))],
    );
  } else if (format === 'docx') {
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    files.push(
      ['word/document.xml', `<w:document xmlns:w="${w}" xmlns:a="${a}" ${r} xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p>${fonts ? '' : `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>${picture.replaceAll('a:blipFill', 'pic:blipFill')}<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`}</w:p><w:sectPr><w:pgSz w:w="3000" w:h="3000"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>`],
      ['word/_rels/document.xml.rels', rels(rel('img', 'image', 'media/limit.png'))],
    );
    if (fonts) files.push(
      ['word/fontTable.xml', `<w:fonts xmlns:w="${w}" ${r}>${fontRefs.map(i => `<w:font w:name="Font ${i}"><w:embedRegular r:id="f${i}" w:fontKey="{00112233-4455-6677-8899-AABBCCDDEEFF}"/></w:font>`).join('')}</w:fonts>`],
      ['word/_rels/fontTable.xml.rels', rels(fontRefs.map(i => rel(`f${i}`, 'font', `fonts/font${i}.ttf`)).join(''))],
    );
  } else {
    files.push(
      ['_rels/.rels', rels(rel('main', 'officeDocument', 'xl/workbook.xml'))],
      ['xl/workbook.xml', `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${r}><sheets><sheet name="S" sheetId="1" r:id="sheet"/></sheets></workbook>`],
      ['xl/_rels/workbook.xml.rels', rels(rel('sheet', 'worksheet', 'worksheets/sheet1.xml'))],
      // SpreadsheetML has font declarations, but no embedded font extraction.
      ['xl/styles.xml', `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3">${fontRefs.map(i => `<font><name val="${fonts && i === 3 ? 'x'.repeat(4097) : `Font ${i}`}"/><sz val="11"/></font>`).join('')}</fonts></styleSheet>`],
      ['xl/worksheets/sheet1.xml', `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ${r}><sheetData/><drawing r:id="drawing"/></worksheet>`],
      ['xl/worksheets/_rels/sheet1.xml.rels', rels(rel('drawing', 'drawing', '../drawings/drawing1.xml'))],
      ['xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${a}" ${r}><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture"/><xdr:cNvPicPr/></xdr:nvPicPr>${picture.replaceAll('a:blipFill', 'xdr:blipFill')}<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`],
      ['xl/drawings/_rels/drawing1.xml.rels', rels(rel('img', 'image', '../media/limit.png'))],
    );
  }
  const prefix = format === 'docx' ? 'word' : format === 'pptx' ? 'ppt' : 'xl';
  if (fonts && format !== 'xlsx') {
    for (const i of fontRefs) files.push([`${prefix}/fonts/font${i}.${format === 'docx' ? 'ttf' : 'fntdata'}`, i === 3 ? 'x'.repeat(4097) : `font${i}`]);
  }
  files.push([`${prefix}/media/limit.png`, 'x'.repeat(4001)]);
  return storedZip(files);
}

class Canvas {
  width = 1;
  height = 1;
  style = {};
  readonly context = new Proxy<Record<string, unknown>>({}, {
    get: (target, key: string) => key in target ? target[key]
      : key === 'measureText' ? () => ({ width: 0, actualBoundingBoxAscent: 0, actualBoundingBoxDescent: 0 })
      : () => undefined,
    set: (target, key: string, value) => { target[key] = value; return true; },
  });
  getContext() { return this.context; }
  transferToImageBitmap() { return { close() {} }; }
}

async function worker(format: Format, entry: 'render-worker' | 'worker' | 'worker-source' = 'render-worker') {
  vi.resetModules();
  const messages: Record<string, unknown>[] = [];
  const fonts = new Set<FontFace>();
  let registeredFaces = 0;
  const scope = {
    fonts: { add: (face: FontFace) => { registeredFaces++; fonts.add(face); }, delete: (face: FontFace) => fonts.delete(face) },
    onmessage: null as null | ((event: MessageEvent) => Promise<void>),
    postMessage: (message: Record<string, unknown>) => { messages.push(message); },
  };
  vi.stubGlobal('self', scope);
  vi.stubGlobal('OffscreenCanvas', Canvas);
  vi.stubGlobal('FontFace', class {
    constructor(readonly family: string, _source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
    get weight() { return this.descriptors.weight; }
    get style() { return this.descriptors.style; }
    load() { return Promise.resolve(this); }
  });
  // Actual worker entry, actual parser, actual renderers; only browser surfaces
  // and font decoding are replaced. No replica of the worker's load logic.
  await workerEntries[format][entry]();
  const dispatch = async (data: unknown) => { await scope.onmessage!({ data } as MessageEvent); };
  const wasm = readFileSync(new URL(`../../${format}/src/wasm/${format}_parser_bg.wasm`, import.meta.url));
  await dispatch({ [format === 'pptx' ? 'kind' : 'type']: 'init', wasmUrl: `data:application/wasm;base64,${wasm.toString('base64')}` });
  return { dispatch, messages, fonts, registeredFaces: () => registeredFaces };
}

const policy: { maxArchiveEntryBytes: number | null; maxTotalInflatedBytes: number | null; maxArchiveEntries: number | null } = { maxArchiveEntryBytes: null, maxTotalInflatedBytes: 4000, maxArchiveEntries: null };
const discriminant = (format: Format) => format === 'pptx' ? 'kind' : 'type';
const parse = (format: Format, data: Uint8Array<ArrayBuffer>, resourcePolicy: typeof policy) => ({
  [discriminant(format)]: 'parse', id: 1, data: data.buffer, buffer: data.buffer,
  defaultCurrentDateMs: 0, resourcePolicy,
});
async function limitResponse(message: Record<string, unknown>, format: Format) {
  const { deserializeWorkerError } = await import('@silurus/ooxml-core/worker');
  const { OoxmlResourceLimitError } = await import('@silurus/ooxml-core');
  expect(message).toMatchObject({ [discriminant(format)]: 'error' });
  const error = deserializeWorkerError(message as never);
  expect(error).toBeInstanceOf(OoxmlResourceLimitError);
  return error as InstanceType<typeof OoxmlResourceLimitError>;
}

afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); vi.doUnmock('@silurus/ooxml-core'); });

describe.each<Format>(['pptx', 'docx', 'xlsx'])('%s real-WASM worker resource boundaries', format => {
  it.each(format === 'pptx' ? ['ordinary', 'progressive', 'empty'] : ['ordinary'])('routes a font crossing through %s load failure without publishing ready or retaining faces', async mode => {
    const { dispatch, messages, fonts, registeredFaces } = await worker(format);
    const data = fixture(format, true);
    // The empty-deck case replaces only sldIdLst; the font list still belongs
    // to the presentation and cannot escape admission because there is no paint.
    const input = mode === 'empty' ? fixture('pptx', true, true) : data;
    await dispatch({ ...parse(format, input, { ...policy, maxArchiveEntryBytes: 4096, maxTotalInflatedBytes: null }), progressiveLayout: mode === 'progressive' });
    expect(messages).toHaveLength(1);
    const error = await limitResponse(messages[0], format);
    expect(error.details.violation).toMatchObject({ resource: 'archive-entry', metric: 'declared-inflated-bytes', limit: 4096 });
    expect(fonts.size).toBe(0);
    if (format === 'pptx') expect(registeredFaces()).toBe(2);
  });

  it('fails the first display-resolution render on a total-inflation image crossing', async () => {
    const { dispatch, messages } = await worker(format);
    await dispatch(parse(format, fixture(format), policy));
    expect(messages[0]).toMatchObject({ [discriminant(format)]: format === 'pptx' ? 'presentationReady' : format === 'docx' ? 'parsedMeta' : 'parsed' });
    if (format === 'xlsx') {
      const { XlsxWorksheetPullClient } = await import('../../xlsx/src/worksheet-pull-client.js');
      let requestId = 10;
      const client = new XlsxWorksheetPullClient({
        sharedStrings: [],
        transport: {
          request: async build => {
            const start = messages.length;
            await dispatch(build(++requestId));
            return messages[start] as never;
          },
          forgetOrphaned() {}, terminate() {},
        },
        open: async (sheetIndex, sheetName, identity) => {
          await dispatch({ type: 'openSheetSession', id: ++requestId, sheetIndex, sheetName, ...identity });
        },
      });
      for await (const unit of client.stream(0, 'S')) {
        if (unit.kind === 'finished') expect(unit.worksheet.images).toHaveLength(1);
      }
    }
    const before = messages.length;
    await dispatch(format === 'pptx'
      ? { kind: 'renderSlide', id: 2, slideIndex: 0, width: 100, dpr: 1, imageResources: { resolution: 'display' } }
      : format === 'docx'
      ? { type: 'renderPage', id: 2, pageIndex: 0, opts: { width: 100, dpr: 1, imageResources: { resolution: 'display' } } }
      : { type: 'renderViewport', id: 2, sheetIndex: 0, viewport: { row: 1, col: 1, rows: 10, cols: 10 }, opts: { width: 100, height: 100, dpr: 1, imageResources: { resolution: 'display' } } });
    expect(messages.slice(before)).toHaveLength(1);
    const error = await limitResponse(messages[before], format);
    expect(error.details.violation).toMatchObject({ resource: 'archive', metric: 'distinct-inflated-bytes', limit: 4000, observed: 4001 });
  });

  it.each(['worker', 'worker-source'] as const)('normalizes direct image extraction in the %s entry', async entry => {
    const { dispatch, messages } = await worker(format, entry);
    await dispatch(parse(format, fixture(format), policy));
    if (format === 'docx') {
      const opened = messages[0];
      await dispatch({ protocol: 'ooxml-pull-v1', kind: 'close', requestId: 3,
        sessionId: opened.sessionId, operationId: opened.operationId, generation: opened.generation });
    }
    const prefix = format === 'docx' ? 'word' : format === 'pptx' ? 'ppt' : 'xl';
    await dispatch({ [discriminant(format)]: 'extractImage', id: 2, path: `${prefix}/media/limit.png` });
    const error = await limitResponse(messages.at(-1)!, format);
    expect(error.details.violation).toMatchObject({ metric: 'distinct-inflated-bytes', limit: 4000, observed: 4001 });
  });
});

// Window/Node callers use an archive proxy rather than the worker host. Verify
// normalization before any consumer catch, without a worker serialization hop.
describe.each<Format>(['pptx', 'docx', 'xlsx'])('%s direct-WASM archive boundaries', format => {
  it.each(['worker-host', 'in-process-proxy'] as const)('normalizes extraction through %s', async boundary => {
    vi.resetModules();
    const wasm = await wasmBindings[format]();
    const wasmModule = new WebAssembly.Module(readFileSync(new URL(`../../${format}/src/wasm/${format}_parser_bg.wasm`, import.meta.url)));
    const { WasmParserHost } = await import('@silurus/ooxml-core/worker');
    const { WasmRuntimeGenerationHost } = await import('@silurus/ooxml-core/internal/wasm-runtime-generation');
    const { OoxmlResourceLimitError } = await import('@silurus/ooxml-core');
    type Archive = { extract_image(path: string): Uint8Array; extract_media?(path: string): Uint8Array; extract_font?(path: string): Uint8Array; free(): void };
    const Archive = 'DocxArchive' in wasm ? wasm.DocxArchive
      : 'PptxArchive' in wasm ? wasm.PptxArchive : wasm.XlsxArchive;
    const prefix = format === 'docx' ? 'word' : format === 'pptx' ? 'ppt' : 'xl';
    const operations = [
      { method: 'extract_image' as const, path: `${prefix}/media/limit.png`, fonts: false, entry: null, total: 4000n, metric: 'distinct-inflated-bytes', observed: 4001 },
      ...(format === 'pptx' ? [
        { method: 'extract_media' as const, path: `${prefix}/media/limit.png`, fonts: false, entry: null, total: 4000n, metric: 'distinct-inflated-bytes', observed: 4001 },
        { method: 'extract_font' as const, path: 'ppt/fonts/font3.fntdata', fonts: true, entry: 4096n, total: null, metric: 'declared-inflated-bytes', observed: 4097 },
      ] : format === 'docx' ? [
        { method: 'extract_image' as const, path: 'word/fonts/font3.ttf', fonts: true, entry: 4096n, total: null, metric: 'declared-inflated-bytes', observed: 4097 },
      ] : []),
    ];
    for (const operation of operations) {
      let extract: () => Uint8Array;
      let close: () => void;
      const create = () => new Archive(fixture(format, operation.fonts), operation.entry, operation.total);
      if (boundary === 'worker-host') {
        const host = new WasmParserHost<Archive>(async () => wasm.initSync({ module: wasmModule }), { freeArchive: archive => archive.free() });
        host.setWasmInput(wasmModule);
        await host.ensureReady();
        host.setArchive(host.run(create));
        extract = () => host.run(() => host.archive![operation.method]!(operation.path));
        close = () => host.disposeArchive();
      } else {
        const host = new WasmRuntimeGenerationHost<Archive>(wasm, wasmModule);
        const handle = await host.open(create);
        extract = () => handle.proxy[operation.method]!(operation.path);
        close = () => handle.close(archive => archive.free());
      }
      try {
        let failure: unknown;
        try { extract(); } catch (error) { failure = error; }
        expect(failure).toBeInstanceOf(OoxmlResourceLimitError);
        expect((failure as InstanceType<typeof OoxmlResourceLimitError>).details.violation).toMatchObject({
          metric: operation.metric, limit: operation.fonts ? 4096 : 4000, observed: operation.observed,
        });
        expect(extract).toThrow(OoxmlResourceLimitError);
      } finally { close(); }
    }
  });
});

it.each(['parse', 'parseDelimitedText'])('XLSX %s awaits asynchronous fonts and recovers after load failure', async type => {
  let rejectFont!: (error: Error) => void;
  const pendingFont = new Promise<never>((_resolve, reject) => { rejectFont = reject; });
  const load = vi.fn().mockReturnValueOnce(pendingFont).mockResolvedValue({ faces: [], routes: {}, checked: [] });
  vi.doMock('@silurus/ooxml-core', async importOriginal => ({
    ...await importOriginal<typeof import('@silurus/ooxml-core')>(), loadOfficeFontFallbacks: load,
  }));
  const { dispatch, messages } = await worker('xlsx');
  const request = type === 'parse' ? parse('xlsx', fixture('xlsx'), policy) : {
    type, id: 1, data: new TextEncoder().encode('a,b\n').buffer,
    options: { delimiter: ',', encoding: 'utf-8', sheetName: 'S' },
  };
  const loading = dispatch(request);
  await vi.waitFor(() => expect(load).toHaveBeenCalledOnce());
  expect(messages).toEqual([]);
  rejectFont(new Error('font load failed'));
  await loading;
  expect(messages).toEqual([expect.objectContaining({ type: 'error', id: 1, message: 'font load failed' })]);
  await dispatch({ ...request, id: 2 });
  expect(messages.at(-1)).toMatchObject({ type: type === 'parse' ? 'parsed' : 'delimitedTextParsed', id: 2 });
});
