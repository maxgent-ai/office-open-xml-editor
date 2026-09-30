import { test, expect } from '@playwright/test';

/**
 * Regression twin of the pptx responsive-overlay spec (same reported bug on
 * https://ooxml.silurus.dev): the docx find-highlight / text-selection overlays
 * placed literal-px boxes sized to the page's INTENDED CSS box over a `<canvas>`
 * a consumer had scaled DOWN with external CSS, so the oversized overlay pushed a
 * scrollbar onto the ancestor scroll area.
 *
 * The fix positions every overlay box as a PERCENTAGE of the intended CSS box and
 * leaves the container at `width:100%;height:100%` so it tracks the canvas's
 * ACTUAL rendered size. This spec asserts the stage never gains a scrollbar and
 * the overlay layers match the scaled canvas — not the intended box.
 */
test('find/selection overlays do not overflow a scaled-canvas scroll area › demo/sample-1', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/tests/visual/responsive-overlay-fixture.html?docx=demo/sample-1&q=the');
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
    overlayRects: { w: number; h: number }[];
  };

  // Precondition: the canvas really was scaled DOWN below its intended 794px box.
  expect(r.canvasCssWidth).toBeLessThan(794);
  expect(r.canvasCssWidth).toBeGreaterThan(0);

  // The reported symptom is a HORIZONTAL scrollbar from the oversized overlay
  // (width 794 while the canvas rendered ~480): the horizontal scroll area must
  // NOT overflow (±1px layout slack). Vertical scrolling is legitimate here — a
  // docx page is genuinely taller than the stage — so it is asserted via the
  // overlay-tracks-canvas check below, which pins BOTH axes of the overlay to the
  // scaled canvas (an over-tall overlay would fail it) without conflating with the
  // page's own height.
  expect(r.stageScrollWidth).toBeLessThanOrEqual(r.stageClientWidth + 1);

  // The overlay layers track the SCALED canvas on BOTH axes, not the intended
  // 794px box — this is the direct fix assertion: an overlay pinned to the
  // intended box (the bug) would be wider AND taller than the scaled canvas.
  // Exactly the text-selection layer and the find-highlight layer: a count
  // check keeps an unrelated wrapper child from standing in for an overlay.
  expect(r.overlayRects.length).toBe(2);
  for (const rect of r.overlayRects) {
    expect(Math.abs(rect.w - r.canvasCssWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(rect.h - r.canvasCssHeight)).toBeLessThanOrEqual(1);
  }
});
