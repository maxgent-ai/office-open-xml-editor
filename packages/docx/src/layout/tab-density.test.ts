/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it, vi } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { documentBytes, measureContext } from '../test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { runLineBreakerPass } from '../line-breaker/pass-driver.js';
import type { LayoutSeg } from '../line-breaker/model.js';
import * as passOperations from '../line-breaker/pass-operations.js';
import * as tabs from '../line-breaker/tabs.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function tabDocument(count: number, rtl: boolean, float: boolean) {
  const archive = new DocxArchive(documentBytes('', rtl, 0,
    { alignment: 'left', count, text: 'word', noFloat: !float }));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
}

function checkLayout(model: ReturnType<typeof tabDocument>) {
  const services = createLayoutServices(model, { measureContext: measureContext() });
  const layout = layoutDocument(model, services, { currentDateMs: 0 });
  const text = layout.pages.flatMap((page) => page.layers.body)
    .filter((node) => node.kind === 'paragraph')
    .flatMap((node) => node.lines.flatMap((line) => line.placements))
    .filter((node) => node.kind === 'text').map((node) => node.text).join('');
  expect(text).toBe('word');
}

// Empty cells beyond the margin can collapse onto one RTL line. Count real
// cell/stop visits: reprojecting each completed prefix would process O(n²)
// cells. The 8k-to-32k ratio must remain below 8 (quadratic growth is 16), and
// each size must independently meet a linear bound. Deterministic work counts
// cover all former stress/scaling cases without JIT, GC or machine-speed noise.
it.each([false, true].flatMap((rtl) => [false, true].map((float) => ({ rtl, float }))))(
  'scales tab work from 8,000 to 32,000 tabs linearly: rtl=$rtl float=$float', ({ rtl, float }) => {
    const small = expectLinearTabWork(8_000, () => checkPass(8_000, rtl, float));
    const large = expectLinearTabWork(32_000, () => checkPass(32_000, rtl, float));
    expect(large / small).toBeLessThan(8);
  }, 120_000,
);

it('resolves parser-backed 32,000-tab RTL cells in linear work', () => {
  const count = 32_000;
  expectLinearTabWork(count, () => checkLayout(tabDocument(count, true, false)));
}, 120_000);

function expectLinearTabWork(count: number, run: () => void) {
  let visits = 0;
  const resolveDecimal = passOperations.performDecimalAlignmentPoint;
  const nextStop = tabs.nextLineTabStop;
  // Delegate to production behavior. The bidi resolver visits every tab cell,
  // including empty cells, through decimal alignment; the LTR iterator selects
  // a stop on each tab attempt. No mock results supply the expected work bound.
  const cells = vi.spyOn(passOperations, 'performDecimalAlignmentPoint')
    .mockImplementation((segments) => {
      visits += 1;
      return resolveDecimal(segments);
    });
  const stops = vi.spyOn(tabs, 'nextLineTabStop')
    .mockImplementation((...args) => {
      visits += 1;
      return nextStop(...args);
    });
  try {
    run();
    expect(visits).toBeGreaterThanOrEqual(count);
    // Completed-cell fitting and final projection each walk once. Allow two
    // more linear walks for retries, rather than any machine-speed allowance.
    expect(visits).toBeLessThan(count * 4);
    return visits;
  } finally {
    cells.mockRestore();
    stops.mockRestore();
  }
}

// Exercise the queue-owning production pass separately from pagination:
// LTR unreachable tabs intentionally produce many lines, and page-level float
// retries have different costs from resolving a single line's completed cells.
function checkPass(count: number, rtl: boolean, float: boolean) {
  const segs: LayoutSeg[] = Array.from({ length: count }, () =>
    ({ isTab: true, fontSize: 12, measuredWidth: 0 }));
  segs.push({ text: 'word', bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 12, color: null, fontFamily: null,
    vertAlign: null, measuredWidth: 0 });
  const lines = runLineBreakerPass({
    ctx: measureContext(), segs, maxWidth: 468, firstIndent: 0, scale: 1,
    tabStops: [], fontFamilyClasses: {}, tabOriginPx: 0,
    kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36, marginRightPx: 468,
    baseRtl: rtl, isJustified: false, stretchLastLine: false,
    widthPolicy: 'bounded', overflowPunct: false,
    passContext: { probeHeights: float ? Array(count + 1).fill(12) : null },
    wrapCtx: float ? {
      startPageY: 0, paraX: 0, columnXPt: 0, columnWidthPt: 468,
      floats: [{ kind: 'shape', mode: 'square', imageKey: '',
        imageX: 0, imageY: 0, imageW: 200, imageH: 100,
        xLeft: 0, xRight: 200, yTop: 0, yBottom: 100, side: 'bothSides',
        distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 }],
      lineBoxH: () => 12, pageH: Number.POSITIVE_INFINITY,
    } : undefined,
  });
  expect(lines.flatMap((line) => line.segments).filter((seg) => 'isTab' in seg)).toHaveLength(count);
  expect(lines.flatMap((line) => line.segments).filter((seg) => 'text' in seg)
    .map((seg) => seg.text).join('')).toBe('word');
}
