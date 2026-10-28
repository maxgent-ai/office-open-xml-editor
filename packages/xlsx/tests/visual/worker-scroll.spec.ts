import { test, expect } from '@playwright/test';

for (const format of ['xlsx', 'csv'] as const) {
  test(`worker ${format} viewer presents frames before continuous scrolling stops`, async ({ page }) => {
    test.setTimeout(60_000);
    await page.route('**/scroll-probe.html', (route) => route.fulfill({
      contentType: 'text/html',
      body: '<html><body style="margin:0"><div id="host" style="width:900px;height:600px"></div></body></html>',
    }));
    await page.goto('/scroll-probe.html');
    const result = await page.evaluate(async (format) => {
      // The route supplies a blank host; initialize Vite's source-mode defines
      // before importing the same entry point consumers use.
      await import('/@vite/env');
      const { XlsxWorkbook, XlsxViewer, XlsxSheetViewer } = await import('/src/index.ts');
      const host = document.getElementById('host') as HTMLDivElement;
      const canvas = document.createElement('canvas');
      canvas.style.width = '900px';
      canvas.style.height = '600px';
      if (format === 'csv') host.appendChild(canvas);
      const viewer = format === 'xlsx'
        ? new XlsxViewer(host, { mode: 'worker' })
        : new XlsxSheetViewer(canvas, { mode: 'worker' });
      const originalRender = XlsxWorkbook.prototype.renderViewportToBitmap;
      const originalTransfer = ImageBitmapRenderingContext.prototype.transferFromImageBitmap;
      let commits = 0;
      let delay = false;
      let beforeImage = '';
      let pixelsChanged = false;
      ImageBitmapRenderingContext.prototype.transferFromImageBitmap = function (bitmap) {
        commits++;
        const result = originalTransfer.call(this, bitmap);
        if (delay) pixelsChanged ||= viewer.canvasElement.toDataURL() !== beforeImage;
        return result;
      };
      XlsxWorkbook.prototype.renderViewportToBitmap = async function (...args: unknown[]) {
        const bitmap = await originalRender.apply(this, args);
        // Keep this regression deterministic on fast machines: a worker
        // response spans multiple scroll events. This is a test delay, not a
        // production throttle or a performance target.
        if (delay) await new Promise(resolve => setTimeout(resolve, 80));
        return bitmap;
      };
      try {
        if (format === 'xlsx') await viewer.load('/demo/sample-1.xlsx');
        else {
          const csv = Array.from({ length: 300 }, (_, row) =>
            Array.from({ length: 20 }, (_, col) => `row ${row + 1} column ${col + 1}`).join(','))
            .join('\n');
          await viewer.load(new TextEncoder().encode(csv).buffer,
            { format: 'csv' });
        }
        const scrollHost = host.querySelector('[data-xlsx-viewport-input]') as HTMLDivElement;
        const paintedCanvas = viewer.canvasElement as HTMLCanvasElement;
        beforeImage = paintedCanvas.toDataURL();
        const before = commits;
        delay = true;
        await new Promise<void>((resolve) => {
          let frame = 0;
          const scroll = () => {
            // Repeat inside the scrollable extent so reaching its end cannot
            // accidentally give the old implementation a quiet interval.
            scrollHost.scrollTop = (++frame % 20) * 10;
            if (frame < 90) requestAnimationFrame(scroll);
            else resolve();
          };
          requestAnimationFrame(scroll);
        });
        return {
          commitsDuringScroll: commits - before,
          pixelsChanged,
          scrollable: scrollHost.scrollHeight > scrollHost.clientHeight + 190,
        };
      } finally {
        viewer.destroy();
        XlsxWorkbook.prototype.renderViewportToBitmap = originalRender;
        ImageBitmapRenderingContext.prototype.transferFromImageBitmap = originalTransfer;
      }
    }, format);
    expect(result.scrollable).toBe(true);
    expect(result.commitsDuringScroll).toBeGreaterThan(0);
    expect(result.pixelsChanged, JSON.stringify(result)).toBe(true);
  });
}
