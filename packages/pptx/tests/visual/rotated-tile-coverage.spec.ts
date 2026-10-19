import { test, expect } from '@playwright/test';

test('nonrotating tiles cover the complete non-square shape with rotation and reflection', async ({ page }) => {
  await page.route('**/tile-probe.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html>' }));
  await page.goto('/tile-probe.html');
  const results = await page.evaluate(async () => {
    const { paintPreparedShapeImageFill } = await import('/src/renderer.ts') as typeof import('../../src/renderer.js');
    const tile = new OffscreenCanvas(4, 2);
    const tileCtx = tile.getContext('2d')!;
    tileCtx.fillStyle = '#FF0000';
    tileCtx.fillRect(0, 0, 4, 2);
    tileCtx.fillStyle = '#0000FF';
    tileCtx.fillRect(2, 0, 2, 2);
    const bitmap = await createImageBitmap(tile);
    const results = [];
    for (const rotation of [0, 45, 90]) {
      for (const flipH of [false, true]) {
        const render = (tiled: boolean, angle = rotation) => {
          const canvas = document.createElement('canvas');
          canvas.width = canvas.height = 240;
          const ctx = canvas.getContext('2d')!;
          ctx.imageSmoothingEnabled = false;
          ctx.translate(120, 120);
          ctx.rotate(angle * Math.PI / 180);
          ctx.scale(flipH ? -1 : 1, 1);
          ctx.beginPath();
          ctx.rect(-50, -25, 100, 50);
          if (tiled) {
            paintPreparedShapeImageFill(ctx, {
              fillType: 'image', imagePath: '', mimeType: 'image/png', rotWithShape: false,
              tile: { sx: 1, sy: 1, tx: 0, ty: 0, algn: 'ctr', flip: 'xy' },
            }, { image: bitmap }, { x: -50, y: -25, w: 100, h: 50 }, 1 / 9525, false,
            { rotation: angle, flipH, flipV: false });
          } else {
            ctx.fillStyle = '#FF0000';
            ctx.fill();
          }
          return ctx.getImageData(0, 0, 240, 240).data;
        };
        const mask = render(false);
        const actual = render(true);
        const unrotated = render(true, 0);
        let phaseDifferences = 0;
        let missing = 0;
        let outside = 0;
        for (let i = 3; i < mask.length; i += 4) {
          if (mask[i] === 255 && actual[i] < 254) missing++;
          if (mask[i] === 0 && actual[i] > 0) outside++;
          if (mask[i] === 255 && actual[i] === 255 && unrotated[i] === 255
            && actual[i - 3] !== unrotated[i - 3]) phaseDifferences++;
        }
        results.push({ rotation, flipH, missing, outside, phaseDifferences });
      }
    }
    bitmap.close();
    return results;
  });
  for (const result of results) {
    expect(result, JSON.stringify(result)).toMatchObject({ missing: 0, outside: 0, phaseDifferences: 0 });
  }
});
