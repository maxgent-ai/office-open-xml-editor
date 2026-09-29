import { expect, test } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

const packageRoot = fileURLToPath(new URL('../../packages/pptx/', import.meta.url));
const fixturePath = fileURLToPath(new URL('./fixtures/pptx-field-outline-projection.pptx', import.meta.url));
let server: ViteDevServer;
let url: string;

// This fixture is a small, synthetic OPC package. Both render modes parse its
// actual a:fld, a:ln and a:scene3d XML through the production WASM worker.
test.beforeAll(async () => {
  server = await createServer({
    root: packageRoot,
    configFile: resolve(packageRoot, 'vite.config.ts'),
    server: { host: '127.0.0.1', port: 0, strictPort: false },
  });
  await server.listen();
  const address = server.httpServer?.address();
  if (!address || typeof address === 'string') throw new Error('Vite did not assign a port');
  url = `http://127.0.0.1:${address.port}/tests/visual/pattern-parser-fixture.html`;
});
test.afterAll(async () => { await server?.close(); });

test('field-local pattern, patterned outline, and projected text retain their paints in both render modes', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const result = await page.evaluate(async bytes => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & { renderPatternParserFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered> }).renderPatternParserFixture;
    const count = (image: Rendered, region?: [number, number, number, number]) => {
      const [x0, y0, x1, y1] = region ?? [0, 0, image.width, image.height];
      let red = 0, cyan = 0, green = 0, orange = 0, black = 0;
      for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
        const i = 4 * (y * image.width + x);
        const [r, g, b, a] = image.bytes.slice(i, i + 4);
        if (a < 220) continue;
        if (r > 165 && g < 90 && b > 55 && b < 180) red++;
        if (r < 95 && g > 155 && b > 155) cyan++;
        if (r < 100 && g > 110 && b < 130) green++;
        if (r > 170 && g > 65 && g < 185 && b < 90) orange++;
        if (r < 50 && g < 50 && b < 50) black++;
      }
      return { red, cyan, green, orange, black };
    };
    const out = [];
    const projectedPhase = (image: Rendered) => {
      let visible = 0, opposite = 0;
      const classify = (x: number, y: number) => {
        const i = 4 * (y * image.width + x);
        const [r, g, b, a] = image.bytes.slice(i, i + 4);
        if (a < 220) return 0;
        if (r > 165 && g < 90 && b > 55 && b < 180) return 1;
        if (r < 95 && g > 155 && b > 155) return 2;
        return 0;
      };
      // The identical projected shapes are separated by 6 in + 4 pt. Six
      // inches is an integer number of 8 pt tiles; the remaining 4 pt must
      // shift the dnDiag colours in their common glyph interiors.
      const dx = Math.round((5486400 + 50800) * image.width / 12192000);
      for (let y = 40; y < 240; y++) for (let x = 60; x < 450; x++) {
        const a = classify(x, y), b = classify(x + dx, y);
        if (a && b) { visible++; if (a !== b) opposite++; }
      }
      return { visible, opposite };
    };
    for (const mode of ['main', 'worker']) {
      const field = await render(bytes, 0, mode);
      const outline = await render(bytes, 1, mode);
      const projected = await render(bytes, 2, mode);
      out.push({ mode,
        field: count(field), outline: count(outline),
        projected: count(projected), projectedPhase: projectedPhase(projected),
        fieldTop: count(field, [50, 50, 620, 210]),
        fieldBottom: count(field, [50, 240, 620, 450]),
        fieldLower: count(field, [50, 470, 620, 670]),
      });
    }
    return out;
  }, bytes);
  for (const row of result) {
    expect(row.fieldTop.red, row.mode).toBeGreaterThan(20);
    expect(row.fieldTop.cyan, row.mode).toBeGreaterThan(20);
    expect(row.fieldBottom.red, row.mode).toBeGreaterThan(20);
    expect(row.fieldBottom.cyan, row.mode).toBeGreaterThan(20);
    expect(row.fieldLower.red, row.mode).toBeGreaterThan(20);
    expect(row.fieldLower.cyan, row.mode).toBeGreaterThan(20);
    expect(row.outline.green, row.mode).toBeGreaterThan(20);
    expect(row.outline.orange, row.mode).toBeGreaterThan(20);
    expect(row.projected.red, row.mode).toBeGreaterThan(20);
    expect(row.projected.cyan, row.mode).toBeGreaterThan(20);
    expect(row.projectedPhase.visible, row.mode).toBeGreaterThan(100);
    expect(row.projectedPhase.opposite, row.mode).toBeGreaterThan(100);
  }
  expect(result[0].field).toEqual(result[1].field);
  expect(result[0].outline).toEqual(result[1].outline);
  expect(result[0].projected).toEqual(result[1].projected);
});
