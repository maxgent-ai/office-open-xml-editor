/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { afterEach, beforeAll, expect, it, vi } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { storedZip } from '@silurus/ooxml-core/testing';
import { resourcePolicyForWasm, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { PptxPresentation } from './presentation.js';
import { PresentationPreflightBuilder } from './presentation-preflight.js';
import type { PptxWorkerRequest, PresentationBootstrap } from './worker-protocol.js';
import { initSync, PptxArchive } from './wasm/pptx_parser.js';

vi.mock('./worker.ts?worker&inline', () => ({ default: function () { return new ArchiveWorker(); } }));

// Exercise load + WorkerBridge + real WASM extraction. Only the browser worker
// transport and FontFace decoder are replaced; archive limits are not mocked.
class ArchiveWorker {
  static latest: ArchiveWorker;
  static preflightDelayMs = 0;
  failureAfterPreflight = false;
  terminated = false;
  preflightFinished = false;
  private archive?: PptxArchive;
  private bootstrap?: PresentationBootstrap;
  private listeners = new Set<(event: MessageEvent) => void>();
  constructor() { ArchiveWorker.latest = this; }
  addEventListener(type: string, listener: (event: MessageEvent) => void) {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(_type: string, listener: (event: MessageEvent) => void) { this.listeners.delete(listener); }
  postMessage(request: PptxWorkerRequest) {
    queueMicrotask(() => {
      try {
        if (request.kind === 'parse') {
          this.archive = new PptxArchive(new Uint8Array(request.buffer), ...resourcePolicyForWasm(request.resourcePolicy));
          this.bootstrap = JSON.parse(new TextDecoder().decode(this.archive.presentation_bootstrap()));
          this.reply({ kind: 'presentationOpened', id: request.id, bootstrap: this.bootstrap });
        } else if (request.kind === 'finishPresentationPreflight') {
          const finish = () => {
            this.preflightFinished = true;
            this.reply({ kind: 'presentationPreflightReady', id: request.id,
              preflight: new PresentationPreflightBuilder(this.bootstrap!).finish() });
          };
          if (ArchiveWorker.preflightDelayMs) setTimeout(finish, ArchiveWorker.preflightDelayMs);
          else finish();
        } else if (request.kind === 'extractFont') {
          this.reply({ kind: 'fontExtracted', id: request.id, bytes: this.archive!.extract_font(request.path).buffer });
        }
      } catch (error) {
        this.failureAfterPreflight = this.preflightFinished;
        this.reply({ kind: 'error', id: 'id' in request ? request.id : undefined, ...serializeWorkerError(error) });
      }
    });
  }
  extractFont(path: string) {
    try { return this.archive!.extract_font(path); }
    catch (error) { this.failureAfterPreflight = this.preflightFinished; throw error; }
  }
  private reply(data: unknown) { for (const listener of this.listeners) listener({ data } as MessageEvent); }
  terminate() { this.terminated = true; this.archive?.free(); this.archive = undefined; }
}

beforeAll(() => {
  initSync({ module: readFileSync(new URL('./wasm/pptx_parser_bg.wasm', import.meta.url)) });
});
afterEach(() => { ArchiveWorker.latest?.terminate(); vi.unstubAllGlobals(); vi.restoreAllMocks(); });

it.each([
  { direct: false, delay: 0 },
  { direct: false, delay: 25 },
  { direct: true, delay: 0 },
])('rejects load and releases earlier faces (direct WASM=$direct, preflight delay=$delay)', async ({ direct, delay }) => {
  // Model worker rendering's direct archive callback before error serialization.
  if (direct) vi.spyOn(
    PptxPresentation.prototype as unknown as { getFontBytes(path: string): Promise<Uint8Array> },
    'getFontBytes',
  ).mockImplementation(async path => ArchiveWorker.latest.extractFont(path));
  ArchiveWorker.preflightDelayMs = delay;
  const faces = new Set<FontFace>();
  let registered = 0;
  vi.stubGlobal('document', { fonts: {
    add: (face: FontFace) => { registered++; faces.add(face); },
    delete: (face: FontFace) => faces.delete(face),
  } });
  vi.stubGlobal('FontFace', class {
    constructor(readonly family: string, _source: ArrayBuffer, readonly descriptors: FontFaceDescriptors) {}
    get weight() { return this.descriptors.weight; }
    get style() { return this.descriptors.style; }
    load() { return Promise.resolve(this); }
  });
  vi.stubGlobal('location', { href: 'http://localhost/' });
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const files: Array<[string, string]> = [
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="fntdata" ContentType="application/x-font-ttf"/></Types>'],
    ['ppt/presentation.xml', `<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="${rel}"><p:embeddedFontLst>${[1, 2, 3].map(i => `<p:embeddedFont><p:font typeface="Font ${i}"/><p:regular r:id="f${i}"/></p:embeddedFont>`).join('')}</p:embeddedFontLst><p:sldIdLst/><p:sldSz cx="914400" cy="914400"/></p:presentation>`],
    ['ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2, 3].map(i => `<Relationship Id="f${i}" Type="${rel}/font" Target="fonts/font${i}.fntdata"/>`).join('')}</Relationships>`],
    ['ppt/fonts/font1.fntdata', 'first'], ['ppt/fonts/font2.fntdata', 'second'],
    ['ppt/fonts/font3.fntdata', 'x'.repeat(4097)],
  ];
  const loading = PptxPresentation.load(storedZip(files).buffer, { resourceLimits: { maxArchiveEntryBytes: 4096 } });
  await expect(loading).rejects.toBeInstanceOf(OoxmlResourceLimitError);
  await expect(loading).rejects.toMatchObject({ details: { violation: {
    resource: 'archive-entry', metric: 'declared-inflated-bytes', limit: 4096, observed: 4097,
  } } });
  expect(ArchiveWorker.latest.preflightFinished).toBe(true);
  expect(ArchiveWorker.latest.failureAfterPreflight).toBe(delay === 0);
  expect(registered).toBe(2);
  expect(faces.size).toBe(0);
  expect(ArchiveWorker.latest.terminated).toBe(true);
});
