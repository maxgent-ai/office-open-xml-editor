import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

/**
 * Acceptance test for the slide-6 sp3d bevel ring (private/sample-4.pptx, the
 * tilted-ellipse photo with a 15 pt a:ln beige border, bevelT hardEdge,
 * extrusionH and outerShdw under a perspectiveFront camera).
 *
 * History: across sessions the ellipse rim regressed into a "cut" look — the
 * outer part of the ring went missing at the four apices, so the silhouette
 * read as sliced off by straight lines against the white slide. Root cause:
 * `paintBeveledFlat` / `projectScene3dPaint` rasterised the body + edges into
 * an offscreen EXACTLY the size of the shape's bounding box, so the outer half
 * of the centre-aligned 15 pt border (and the bevel band's shading near the
 * box edge) was clipped wherever the ellipse touches its box — i.e. at every
 * apex. The defect exists at every scale but its pixel size grows with the
 * render scale, which is why small-canvas verifications kept missing it.
 *
 * Ground truth (sample-4.pdf p6): the ring is a CONTINUOUS pale-beige band of
 * the FULL border width on every apex — a soft outer shadow outside it, then
 * the bevel-lit beige band straddling the silhouette, then the photo, with NO
 * white interruption, no straight cut, and no saturated-white lit lip.
 *
 * The .pptx is private (never committed). Like the other private-sample specs
 * this one SKIPS when the fixture is absent so a clean checkout stays green.
 * It asserts pixel structure from the live canvas, not a reference image.
 *
 * The check runs at several render scales (width × deviceScaleFactor) because
 * the original regression was scale-dependent in visibility: dpr1/1280 hides a
 * 10 px loss inside the shadow falloff, dpr2 makes it a visible cut.
 */

const SAMPLE = 'private/sample-4';
const SLIDE_INDEX = 5; // PDF p6, 0-based

// EMU geometry of the deck (ppt/presentation.xml sldSz) and the ring border
// (slide6.xml a:ln w="190500" = 15 pt). Used to derive the expected band width
// in device px at each scale.
const SLIDE_W_EMU = 12192000;
const STROKE_W_EMU = 190500;

const fixturePath = join(
  dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  'public',
  `${SAMPLE}.pptx`,
);

/** width = viewer render width (CSS px), dpr = deviceScaleFactor. */
const SCALES: Array<{ width: number; dpr: number }> = [
  { width: 1280, dpr: 1 }, // the VRT default
  { width: 1280, dpr: 2 }, // HiDPI at the same layout width
  { width: 960, dpr: 2 },  // the Storybook card scale the user reproduced at
];

test.describe('pptx slide-6 sp3d bevel ring', () => {
  for (const { width, dpr } of SCALES) {
    test(`full-width continuous beige band on every apex (width=${width}, dpr=${dpr})`, async ({
      browser,
    }, testInfo) => {
      testInfo.skip(!existsSync(fixturePath), `${SAMPLE}.pptx not present (private sample)`);

      const context = await browser.newContext({
        viewport: { width: 1400, height: 1100 },
        deviceScaleFactor: dpr,
      });
      const page = await context.newPage();
      try {
        await page.goto(
          `/tests/visual/fixture.html?pptx=${SAMPLE}&slide=${SLIDE_INDEX}&width=${width}`,
        );
        await page.waitForFunction(
          () =>
            document.body.dataset.status === 'ready' ||
            document.body.dataset.status === 'error',
          { timeout: 30_000 },
        );
        const status = await page.evaluate(() => document.body.dataset.status);
        if (status === 'error') {
          const msg = await page.evaluate(() => document.body.dataset.errorMessage ?? '');
          throw new Error(`Fixture error: ${msg}`);
        }
        // Sample the four apex cross-sections off the live canvas and classify
        // each ray's pixels into page-white / warm-rim / photo. Acceptance
        // contract, per the PDF ground truth:
        //   1. ZERO white pixels inside the rim (no white cut band);
        //   2. a CONTINUOUS warm-toned run of (almost) the FULL border width —
        //      the centre-aligned 15 pt stroke must not lose its outer half to
        //      the offscreen box clip at the apices;
        //   3. the rim (including the bevel-lit lip) stays below page white.
        //
        // The rim classifier is HUE-based (warm: r ≥ b, near-neutral r vs g),
        // not brightness-based: the bevel lip legitimately brightens the lit
        // side and darkens the shadowed side of the beige border, and the
        // exact shading level is a light-rig calibration concern (tracked
        // separately — our threePt rig currently lacks the fill light and the
        // <a:rot rev> rotation, so the bottom lip runs darker than the PDF).
        // What this spec locks is the band's WIDTH and CONTINUITY, which is
        // what the offscreen box clip destroyed.
        //
        // The floor is 0.8 × the border's device-px width: the missing ~0.2
        // allows the antialiased blend into the outer shadow and the photo.
        // The pre-fix renders measure ~0.5 × (half the band clipped), so the
        // floor separates the two states with margin on both sides.
        const result = await page.evaluate(
          ({ slideWEmu, strokeWEmu }) => {
            const canvas = document.querySelector('canvas') as HTMLCanvasElement;
            const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
            const cw = canvas.width;
            const ch = canvas.height;
            const strokePx = (strokeWEmu / slideWEmu) * cw;
            const minRim = Math.max(8, Math.round(strokePx * 0.8));
            const D = ctx.getImageData(0, 0, cw, ch).data;
            const at = (x: number, y: number): [number, number, number] => {
              const o = (Math.round(y) * cw + Math.round(x)) * 4;
              return [D[o], D[o + 1], D[o + 2]];
            };
            const isWhite = (r: number, g: number, b: number) =>
              r >= 250 && g >= 250 && b >= 250;
            const isPhoto = (r: number, g: number, b: number) =>
              b > 95 && b > r + 25 && b > g + 15;
            // Warm rim tone: the C8C6BD border under any bevel lighting (lit
            // ~+25% through shadowed ~-30%), excluding the page white, the
            // neutral shadow greys (r === b) and the blue photo.
            const isRim = (r: number, g: number, b: number) =>
              r > 120 && g > 110 && b > 100 && Math.abs(r - g) < 30 && r > b && !isWhite(r, g, b);

            function analyze(name: string, x0: number, y0: number, dx: number, dy: number) {
              let x = x0;
              let y = y0;
              // Skip leading page white.
              while (x >= 0 && y >= 0 && x < cw && y < ch) {
                const [r, g, b] = at(x, y);
                if (!isWhite(r, g, b)) break;
                x += dx;
                y += dy;
              }
              // March through the rim until the photo, tracking white
              // intrusions and the longest beige run + its brightest pixel.
              // The march is bounded to a few border-widths so a ray that
              // enters dark photo content (the bottom caption) cannot wander
              // deep into the bitmap and pick up unrelated beige runs.
              let whiteInRim = 0;
              let rimRun = 0;
              let bestRimRun = 0;
              let rimMax = 0;
              const maxMarch = Math.ceil(strokePx * 6);
              for (let i = 0; i < maxMarch; i++) {
                if (x < 0 || y < 0 || x >= cw || y >= ch) break;
                const [r, g, b] = at(x, y);
                if (isPhoto(r, g, b)) break;
                if (isWhite(r, g, b)) whiteInRim++;
                if (isRim(r, g, b)) {
                  rimRun++;
                  bestRimRun = Math.max(bestRimRun, rimRun);
                  rimMax = Math.max(rimMax, 0.299 * r + 0.587 * g + 0.114 * b);
                } else {
                  rimRun = 0;
                }
                x += dx;
                y += dy;
              }
              return { name, whiteInRim, bestRimRun, rimMax: Math.round(rimMax) };
            }

            const cx = Math.round(cw / 2);
            const cy = Math.round(ch / 2);
            return {
              minRim,
              strokePx: Math.round(strokePx),
              cw,
              sections: [
                analyze('TOP', cx, 0, 0, 1),
                analyze('BOTTOM', cx, ch - 1, 0, -1),
                analyze('LEFT', 0, cy, 1, 0),
                analyze('RIGHT', cw - 1, cy, -1, 0),
              ],
            };
          },
          { slideWEmu: SLIDE_W_EMU, strokeWEmu: STROKE_W_EMU },
        );

        console.log(
          `slide-6 ring (canvas ${result.cw}px, stroke ${result.strokePx}px, minRim ${result.minRim}px):`,
          JSON.stringify(result.sections),
        );
        for (const r of result.sections) {
          // 1) No white band inside the rim.
          expect(r.whiteInRim, `${r.name}: white pixels inside the ring`).toBe(0);
          // 2) A continuous beige band of (almost) the full border width.
          expect(
            r.bestRimRun,
            `${r.name}: continuous beige run (>= ${result.minRim} of ${result.strokePx} device px)`,
          ).toBeGreaterThanOrEqual(result.minRim);
          // 3) The rim (incl. any lit lip) stays a clear step below page white.
          expect(r.rimMax, `${r.name}: rim brightness must stay below white`).toBeLessThan(250);
        }
      } finally {
        await context.close();
      }
    });
  }
});
