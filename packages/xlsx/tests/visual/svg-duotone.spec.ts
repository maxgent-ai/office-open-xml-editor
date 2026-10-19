import { test, expect } from '@playwright/test';

// Exercise the real Window SVG rasterizer: Chromium cannot reliably decode an
// SVG Blob with createImageBitmap, so a fake bitmap alone misses this failure.
test('SVG-only pictures retain duotone and alpha through browser rasterization', async ({ page }) => {
  await page.goto('/tests/visual/fixture.html?file=demo/sample-1.xlsx&sheet=0');
  const pixels = await page.evaluate(async () => {
    const { decodeImageSource } = await import('/src/render-orchestrator.ts');
    const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="3" height="1">'
      + '<path fill="black" d="M0 0h1v1H0z"/>'
      + '<path fill="white" d="M1 0h1v1H1z"/>'
      + '<path fill="white" fill-opacity="0.5" d="M2 0h1v1H2z"/></svg>';
    const fetchImage = async () => new Blob([svg], { type: 'image/svg+xml' });
    const source = await decodeImageSource(
      'xl/media/duotone.svg', 'image/svg+xml', undefined, fetchImage,
      3, 1, null, { clr1: '000000', clr2: 'FF0000' }, undefined, true,
    );
    if (!source) throw new Error('SVG duotone picture was omitted');
    const canvas = new OffscreenCanvas(3, 1);
    const context = canvas.getContext('2d');
    if (!context) throw new Error('Missing pixel readback context');
    context.drawImage(source, 0, 0);
    return Array.from(context.getImageData(0, 0, 3, 1).data);
  });
  expect(pixels).toEqual([0, 0, 0, 255, 255, 0, 0, 255, 255, 0, 0, 128]);
});

// Exercise the same effect adapter with the real Window rasterizer and the
// worker bridge, including an aspect-mismatched full-source crop target.
test('SVG duotone rasterizes at each display target in Window and worker bridge decoding', async ({ page }) => {
  await page.goto('/tests/visual/fixture.html?file=demo/sample-1.xlsx&sheet=0');
  const results = await page.evaluate(async () => {
    const { decodeImageSource } = await import('/src/render-orchestrator.ts');
    const { WorkerSvgDecodeClient, respondToWorkerSvgDecodeRequest } = await import('/tests/visual/svg-decode-bridge-fixture.ts');
    const results = [];
    for (const mode of ['window', 'worker', 'viewBox-worker']) {
      const worker = mode !== 'window';
      const client = new WorkerSvgDecodeClient((message: unknown) => {
        respondToWorkerSvgDecodeRequest((response: unknown) => client.accept(response), message);
      });
      const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="3" height="1" viewBox="0 0 3 1">'
        + '<path fill="white" d="M0 0h3v1H0z"/>'
        + '<path fill="black" d="M1.49 0h.02v1h-.02z"/></svg>';
      const sourceSvg = mode === 'viewBox-worker' ? svg.replace('width="3" height="1" ', '') : svg;
      const fetchImage = async () => new Blob([sourceSvg], { type: 'image/svg+xml' });
      const grids: number[][] = [];
      const effectSurface = (width: number, height: number) => {
        grids.push([width, height]);
        return new OffscreenCanvas(width, height);
      };
      const sources = [];
      for (const target of [
        { targetWidthPx: 300, targetHeightPx: 100 },
        { targetWidthPx: 600, targetHeightPx: 200 },
        { targetWidthPx: 300, targetHeightPx: 100 },
        { targetWidthPx: 900, targetHeightPx: 100 },
      ]) {
        const source = await decodeImageSource(
          'xl/media/detail.svg', 'image/svg+xml', undefined, fetchImage,
          225, 75, { l: 0.5, t: 0, r: 0, b: 0 }, { clr1: '000000', clr2: 'FF0000' },
          effectSurface, true, undefined, target, worker ? client.decode : undefined,
        );
        if (!source) throw new Error('SVG duotone picture was omitted');
        sources.push(source);
      }
      const source = sources[1] as ImageBitmap;
      const canvas = new OffscreenCanvas(600, 200);
      const context = canvas.getContext('2d');
      if (!context) throw new Error('Missing pixel readback context');
      context.drawImage(source, 0, 0);
      results.push({
        sizes: sources.map((source) => [(source as ImageBitmap).width, (source as ImageBitmap).height]),
        reused: sources[0] === sources[2],
        grids,
        detail: Array.from(context.getImageData(299, 100, 1, 1).data),
        background: Array.from(context.getImageData(100, 100, 1, 1).data),
      });
      client.dispose();
    }
    return results;
  });
  for (const result of results) {
    expect(result.sizes).toEqual([[300, 100], [600, 200], [300, 100], [900, 100]]);
    expect(result.reused).toBe(true);
    expect(result.grids).toEqual([[300, 100], [600, 200], [900, 300]]);
    expect(result.detail).toEqual([0, 0, 0, 255]);
    expect(result.background).toEqual([255, 0, 0, 255]);
  }
});
