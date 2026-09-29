import { expect, test } from '@playwright/test';

async function expectWorkerBitmaps(page: import('@playwright/test').Page, url: string) {
  const sourceRequests: string[] = [];
  const sourceBodies: string[] = [];
  const inspectedResponses: Promise<void>[] = [];
  page.context().on('request', (request) => {
    const pathname = new URL(request.url()).pathname;
    if ((pathname.startsWith('/dist/') || pathname.startsWith('/consumer/'))
      && /model-source|source-worker|worker-source|render-worker-source/.test(pathname)) {
      sourceRequests.push(pathname);
    }
  });
  page.context().on('response', (response) => {
    const pathname = new URL(response.url()).pathname;
    // The consumer fixture embeds a protocol marker in its test source, so
    // inspect the published dist bodies, where a marker means runtime code.
    if (!pathname.startsWith('/dist/') || !/\.m?js$/.test(pathname)) return;
    inspectedResponses.push(response.text().then((body) => {
      if (body.includes('ooxml-model-source-module/v1')
        || body.includes('model source view default')) sourceBodies.push(pathname);
    }));
  });
  await page.goto(`${url}${url.includes('?') ? '&' : '?'}pause-sources`);
  await expect.poll(async () => {
    const body = page.locator('body');
    return (await body.getAttribute('data-ordinary-ready')) === 'true'
      || (await body.getAttribute('data-status')) === 'error';
  }, { timeout: 60_000 }).toBe(true);
  await Promise.all(inspectedResponses);
  expect(sourceRequests, 'ordinary OOXML main/worker loads must not fetch source chunks').toEqual([]);
  expect(sourceBodies, 'ordinary OOXML main/worker loads must not fetch source module bodies').toEqual([]);
  await expect(page.locator('body')).toHaveAttribute('data-ordinary-ready', 'true');
  await expect(page.locator('body')).toHaveAttribute('data-ordinary-loads',
    'docx-worker,xlsx-worker,pptx-worker,docx-main,xlsx-main,pptx-main');
  await page.evaluate(() => (window as typeof window & { resumeSourceStages?: () => void }).resumeSourceStages?.());
  await expect.poll(
    () => page.locator('body').getAttribute('data-status'),
    { timeout: 60_000 },
  ).not.toBe('loading');
  const status = await page.locator('body').getAttribute('data-status');
  if (status !== 'ready') {
    throw new Error(await page.locator('body').getAttribute('data-error-message') ?? status ?? '');
  }

  for (const id of [
    'docx',
    'math',
    'xlsx',
    'pptx',
    'pptx-text',
    'xlsx-bordered',
    'xlsx-csv-main',
    'xlsx-csv-worker',
  ]) {
    const ink = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      // Worker-backed viewers own a `bitmaprenderer` context, so acquiring a
      // second 2D context on their canvas correctly returns null. Copy the
      // displayed bitmap into a disposable 2D canvas before inspecting pixels.
      const sample = document.createElement('canvas');
      sample.width = canvas.width;
      sample.height = canvas.height;
      const context = sample.getContext('2d');
      if (!context) return 0;
      context.drawImage(canvas, 0, 0);
      const pixels = context.getImageData(0, 0, sample.width, sample.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        if (pixels[offset] < 250 || pixels[offset + 1] < 250 || pixels[offset + 2] < 250) count++;
      }
      return count;
    });
    expect(ink, `${id} worker bitmap should contain ink`).toBeGreaterThan(100);
  }

  await expect(page.locator('body')).toHaveAttribute('data-model-sources', 'ready');

  const pptxTextRuns = await page.evaluate(() => (
    window as typeof window & { pptxTextRuns?: Array<Record<string, unknown>> }
  ).pptxTextRuns);
  expect(pptxTextRuns).toHaveLength(1);
  for (const field of ['inShapeX', 'inShapeY', 'w', 'h', 'fontSize']) {
    expect(
      Number.isFinite(pptxTextRuns?.[0]?.[field]),
      `PPTX worker text run ${field} should be finite`,
    ).toBe(true);
  }

  for (const id of ['docx-chart-ex', 'xlsx-chart-ex', 'pptx-chart-ex']) {
    const coloredInk = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      const context = canvas.getContext('2d');
      if (!context) return 0;
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        const red = pixels[offset];
        const green = pixels[offset + 1];
        const blue = pixels[offset + 2];
        if (Math.max(red, green, blue) - Math.min(red, green, blue) > 30) count++;
      }
      return count;
    });
    expect(coloredInk, `${id} worker bitmap should execute the ChartEx painter`)
      .toBeGreaterThan(100);
  }
}

test('published dist starts all three render workers with optional renderers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/');
});

test('Vite consumer bundle preserves all three render workers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/consumer/index.html');
});
