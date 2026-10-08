import { test, expect } from '@playwright/test';

/**
 * Regression for the reported "Single viewer with navigation" scrollbar bug on
 * https://ooxml.silurus.dev/pptx/ : the find-highlight / text-selection overlays
 * placed literal-px boxes sized to the slide's INTENDED CSS box (e.g. 960×540)
 * over a `<canvas>` that a consumer had scaled DOWN with external CSS
 * (`width:100%!important; height:auto`). The oversized overlay overflowed the
 * viewer's wrapper and pushed a scrollbar onto the ancestor scroll area
 * (`scrollWidth`/`scrollHeight` > `clientWidth`/`clientHeight`).
 *
 * The fix positions every overlay box as a PERCENTAGE of the intended CSS box and
 * leaves the overlay container at `width:100%;height:100%`, so it tracks the
 * canvas's ACTUAL rendered size. This spec loads a real `PptxViewer` with a find
 * highlight + selection overlay inside an `overflow:auto` stage that scales the
 * canvas down, and asserts the stage never gains a scrollbar and the overlay
 * layers match the scaled canvas — not the 960×540 intended box.
 *
 * Runs in a real browser (jsdom cannot lay out `%` against a scaled ancestor), so
 * it catches exactly what the unit tests cannot: the resolved geometry.
 */
test('find/selection overlays do not overflow a scaled-canvas scroll area › demo/sample-1', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/tests/visual/responsive-overlay-fixture.html?pptx=demo/sample-1&q=a');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 90_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }

  const r = JSON.parse(await page.evaluate(() => document.body.dataset.result ?? '{}')) as {
    stageScrollWidth: number;
    stageClientWidth: number;
    stageScrollHeight: number;
    stageClientHeight: number;
    canvasCssWidth: number;
    canvasCssHeight: number;
    overlayRects: { tag: string; w: number; h: number }[];
  };

  // Precondition: the canvas really was scaled DOWN below its intended 960×540
  // box (otherwise the test would not exercise the bug at all).
  expect(r.canvasCssWidth).toBeLessThan(960);
  expect(r.canvasCssWidth).toBeGreaterThan(0);

  // The core assertion (the reported symptom): the scroll area does NOT overflow.
  // Before the fix, stageScrollWidth ≈ 960 + padding while clientWidth ≈ 640, so
  // a horizontal scrollbar appeared; likewise vertically. Allow a 1px slack for
  // sub-pixel rounding in the browser's layout.
  expect(r.stageScrollWidth).toBeLessThanOrEqual(r.stageClientWidth + 1);
  expect(r.stageScrollHeight).toBeLessThanOrEqual(r.stageClientHeight + 1);

  // And the overlay layers track the SCALED canvas, not the intended 960×540 box:
  // each overlay layer's laid-out width equals the canvas's rendered width (±1).
  expect(r.overlayRects.map((rect) => rect.tag).sort()).toEqual(['element', 'highlight', 'text']);
  for (const rect of r.overlayRects) {
    expect(Math.abs(rect.w - r.canvasCssWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(rect.h - r.canvasCssHeight)).toBeLessThanOrEqual(1);
  }
});
