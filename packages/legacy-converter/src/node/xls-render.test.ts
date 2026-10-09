// A direct XLS workbook painted by the ordinary XLSX canvas renderer in Node.
import { expect, it } from 'vitest';
import { testXlsSource } from '../test-sources.js';
import {
  installImageBitmapShim, installOffscreenCanvasShim, materializeXlsxWorkbook,
  renderWorksheetViewport, skia, skiaFactory,
} from './node-facade.js';
import { buildXlsRichFixture } from './xls-rich-fixture.js';
import { concat, little16, little32 } from '../test-fixtures.js';
import { workbookCfb } from '../xls-workbook-fixture.js';
import { threeD } from '../../../../src/three-d.js';

it.skipIf(!skia)('paints direct XLS cell text and its rich run color through the XLSX canvas path', async () => {
  const factory = skiaFactory();
  const parsed = await materializeXlsxWorkbook(buildXlsRichFixture(), { modelSources: [testXlsSource()], factory });
  const canvas = factory.createCanvas(700, 200);
  const restoreImage = installImageBitmapShim(factory);
  const restoreOffscreen = installOffscreenCanvasShim(factory);
  try {
    await renderWorksheetViewport({ ws: parsed.worksheets[0]!, styles: parsed.workbookIndex.styles },
      canvas as unknown as HTMLCanvasElement, { row: 1, col: 1, rows: 6, cols: 10 }, { width: 700, height: 200, dpr: 1 });
    const context = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    const pixels = context.getImageData(0, 0, 700, 200).data;
    let red = 0;
    let dark = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      const [r, g, b] = [pixels[i]!, pixels[i + 1]!, pixels[i + 2]!];
      if (r > 220 && g < 30 && b < 30) red++;
      else if (r < 100 && g < 100 && b < 100) dark++;
    }
    // "RED " paints in its run color; the default-colored runs and the
    // number paint dark text.
    expect(red).toBeGreaterThan(0);
    expect(dark).toBeGreaterThan(0);
  } finally { restoreOffscreen(); restoreImage(); }
});

/** Synthetic BIFF chart sheet: one cylinder series with three cached points. */
function cylinderChartSheet(): Uint8Array {
  const words = (...values: number[]) => concat(...values.map(little16));
  const record = (kind: number, data: Uint8Array = new Uint8Array()) => concat(words(kind, data.length), data);
  const bof = (kind: number) => record(0x0809, words(0x0600, kind, 0, 0));
  const number = (row: number, value: number) => {
    const data = new Uint8Array(14); data.set(words(row, 0, 0));
    new DataView(data.buffer).setFloat64(6, value, true); return record(0x0203, data);
  };
  const bound = concat(little32(0), Uint8Array.of(0, 2, 1, 0, 83));
  const globals = () => concat(bof(5), record(0x0085, bound), record(0x000a));
  bound.set(little32(globals().length));
  return workbookCfb(concat(globals(), bof(0x20),
    record(0x1002, concat(little32(0), little32(0), little32(500 * 65536), little32(300 * 65536))),
    record(0x1033), record(0x1003, words(3, 1, 3, 3, 1, 0)), record(0x1033),
    record(0x1006, words(0xffff, 0, 0, 0)), record(0x1033),
    record(0x105f, Uint8Array.of(1, 0)), record(0x1034),
    record(0x1045, words(0)), record(0x1034),
    record(0x1041, new Uint8Array(18)), record(0x1033),
    record(0x1014, new Uint8Array(20)), record(0x1033), record(0x1017, words(0, 150, 0)),
    record(0x103a, words(20, 15, 30, 100, 100, 150, 0x17)),
    record(0x1034), record(0x1034), record(0x1034),
    record(0x1065, words(1)), number(0, 4), number(1, 7), number(2, 2), record(0x000a),
  ));
}

it.skipIf(!skia)('paints a native XLS chart sheet in 3D when the optional painter is supplied', async () => {
  const factory = skiaFactory();
  const parsed = await materializeXlsxWorkbook(cylinderChartSheet(), { modelSources: [testXlsSource()], factory });
  const ws = parsed.worksheets[0]!;
  expect(ws.isChartSheet).toBe(true);
  const flat = factory.createCanvas(700, 390), projected = factory.createCanvas(700, 390);
  const restore = installOffscreenCanvasShim(factory);
  try {
    const model = { ws, styles: parsed.workbookIndex.styles };
    const range = { row: 1, col: 1, rows: 20, cols: 10 };
    await renderWorksheetViewport(model, flat as unknown as HTMLCanvasElement, range, { width: 700, height: 390, dpr: 1 });
    await renderWorksheetViewport({ ...model, threeD }, projected as unknown as HTMLCanvasElement, range, { width: 700, height: 390, dpr: 1 });
    const pixels = (canvas: typeof flat) => Buffer.from((canvas.getContext('2d') as unknown as CanvasRenderingContext2D).getImageData(0, 0, 700, 390).data);
    // The opt-in API must reach native chart-sheet rendering; losing the
    // Chart3d projection or dropping the injected painter makes these equal.
    expect(pixels(projected).equals(pixels(flat))).toBe(false);
  } finally { restore(); }
});
