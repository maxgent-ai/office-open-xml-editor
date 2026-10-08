import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import { tableImageFillBytes } from './table-image-fill-fixtures.mjs';

test.afterEach(async ({ page }) => {
  await page.evaluate(() => {
    (window as unknown as { destroyPptxWorkerVrt?: () => void }).destroyPptxWorkerVrt?.();
  });
});

for (const [kind, top, bottom] of [
  ['theme-background', [102, 0, 153, 255], [0, 0, 255, 255]],
  ['direct-background', [128, 127, 0, 255], [0, 255, 0, 255]],
  ['theme-band', [0, 0, 255, 255], [255, 255, 255, 255]],
  ['direct-band', [0, 255, 0, 255], [255, 255, 255, 255]],
  ['direct-table', [255, 127, 0, 255], [255, 255, 0, 255]],
  ['direct-cell', [255, 255, 0, 255], [255, 255, 0, 255]],
] as const) {
  test(`table image fills: ${kind} parses, composites and matches the worker`, async ({ page }) => {
    await page.route('**/table-image-test.pptx', route => route.fulfill({
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      body: tableImageFillBytes(kind),
    }));
    await page.goto('/tests/visual/worker-fixture.html?pptx=table-image-test&slide=0');
    await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status ?? ''),
      undefined, { timeout: 60_000 });
    expect(await page.evaluate(() => document.body.dataset.errorMessage)).toBeUndefined();
    const urls = await page.evaluate(() => ['main-canvas', 'worker-canvas'].map(id =>
      (document.getElementById(id) as HTMLCanvasElement).toDataURL()));
    const [main, worker] = urls.map(url => PNG.sync.read(Buffer.from(url.split(',')[1], 'base64')));
    expect(worker.width).toBe(main.width);
    expect(worker.height).toBe(main.height);
    expect(worker.data.equals(main.data)).toBe(true);
    const pixel = (y: number) => Array.from(main.data.subarray((y * main.width + 100) * 4,
      (y * main.width + 100) * 4 + 4));
    expect(pixel(100)).toEqual(top);
    expect(pixel(600)).toEqual(bottom);
  });
}

for (const kind of ['pattern-background', 'pattern-cell']) {
  test(`table pattern fills: ${kind} retains both colors in main and worker`, async ({ page }) => {
    await page.route('**/table-pattern-test.pptx', route => route.fulfill({ body: tableImageFillBytes(kind) }));
    await page.goto('/tests/visual/worker-fixture.html?pptx=table-pattern-test&slide=0');
    await page.waitForFunction(() => ['ready', 'error'].includes(document.body.dataset.status ?? ''),
      undefined, { timeout: 60_000 });
    expect(await page.evaluate(() => document.body.dataset.errorMessage)).toBeUndefined();
    const [main, worker] = await page.evaluate(() => ['main-canvas', 'worker-canvas'].map(id =>
      (document.getElementById(id) as HTMLCanvasElement).toDataURL()));
    const a = PNG.sync.read(Buffer.from(main.split(',')[1], 'base64'));
    const b = PNG.sync.read(Buffer.from(worker.split(',')[1], 'base64'));
    expect(b.data.equals(a.data)).toBe(true);
    const colors = new Set<number>();
    for (let y = 500; y < 520; y++) {
      for (let x = 100; x < 120; x++) colors.add(a.data[(y * a.width + x) * 4]);
    }
    expect(colors.has(0)).toBe(true);
    expect(colors.has(255)).toBe(true);
  });
}
