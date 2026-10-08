import { test, expect } from '@playwright/test';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
// Public deterministic synthetic package generator shared with Node tests.
import { generateSyntheticXlsx } from '../../../node/scripts/generate-synthetic-xlsx.mjs';

let directory = '';
const fixtures = new Map<string, Buffer>();

test.beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'xlsx-finite-browser-'));
  for (const [kind, rows, columns] of [['rows', 100_001, 1], ['cells', 12_550, 20]] as const) {
    const path = join(directory, `${kind}.xlsx`);
    await generateSyntheticXlsx(path, { rows, columns });
    fixtures.set(kind, await readFile(path));
  }
});
test.afterAll(async () => { if (directory) await rm(directory, { recursive: true, force: true }); });

for (const kind of ['rows', 'cells'] as const) {
  for (const mode of ['main', 'worker'] as const) {
    for (const override of [false, true]) {
      test(`${kind} ${mode}: ${override ? 'finite override paints, scrolls, finds and copies the tail' : 'default refuses materialization'}`, async ({ page, context }, testInfo) => {
        test.setTimeout(90_000);
        await context.grantPermissions(['clipboard-read', 'clipboard-write']);
        await page.route('**/finite-*.xlsx', async (route) => {
          const bytes = fixtures.get(kind);
          if (!bytes) throw new Error('missing synthetic fixture');
          await route.fulfill({ status: 200, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', body: bytes });
        });
        await page.goto('/tests/visual/finite-policy-fixture.html');
        await page.waitForFunction(() => typeof (window as unknown as { runFinitePolicyCase?: unknown }).runFinitePolicyCase === 'function');
        const result = await page.evaluate(async (args) => {
          const fixture = window as unknown as { runFinitePolicyCase: (input: typeof args) => Promise<Record<string, unknown>> };
          return await fixture.runFinitePolicyCase(args);
        }, { mode, kind, override });
        if (!override) {
          expect(result).toMatchObject({ status: 'refused', code: 'ooxml-resource-limit', details: {
            violation: { metric: kind, limit: kind === 'rows' ? 100_000 : 250_000, configurable: true },
          } });
        } else {
          expect(result).toMatchObject({ status: 'ok', rows: kind === 'rows' ? 100_001 : 12_550,
            cells: kind === 'rows' ? 100_001 : 251_000, copy: { status: 'copied', cellCount: 1 } });
          const tail = kind === 'rows' ? 'A100001' : 'T12550';
          expect(result.matchLocations).toEqual(expect.arrayContaining([expect.objectContaining({ ref: tail })]));
          expect(result.selection).toMatchObject({ kind: 'range', cells: [expect.objectContaining({ displayText: result.tailText })] });
          expect(await page.evaluate(() => navigator.clipboard.readText())).toBe(result.tailText);
          expect((result.offset as { y: number }).y).toBeGreaterThan(0);
          expect(result.resolvedLimits).toMatchObject(kind === 'rows' ? { maxRows: 110_000 } : { maxCells: 300_000 });
          expect(String(result.image)).toMatch(/^data:image\/png;base64,/);
          await page.screenshot({ path: testInfo.outputPath('tail.png') });
        }
        delete result.image;
        const evidencePath = testInfo.outputPath('evidence.json');
        await writeFile(evidencePath, JSON.stringify({ mode, kind, override, ...result }, null, 2));
        await testInfo.attach('finite-policy-evidence', { path: evidencePath, contentType: 'application/json' });
      });
    }
  }
}
