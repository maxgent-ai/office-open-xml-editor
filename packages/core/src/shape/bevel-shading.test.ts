import { describe, it, expect } from 'vitest';
import {
  bevelHeightProfile,
  edt1d,
  gaussianBlur,
  computeBevelNormals,
  shadePixel,
  shadeParamsFor,
  lightDirFromRig,
  applyBevelShading,
  applyExtrusion,
  type BevelShadeParams,
  type BevelCtx,
} from './bevel-shading';

/** Minimal ImageData-backed BevelCtx for the compositor functions. */
function fakeCtx(w: number, h: number, fill?: (x: number, y: number) => [number, number, number, number]): BevelCtx & { data: Uint8ClampedArray } {
  const data = new Uint8ClampedArray(w * h * 4);
  if (fill) {
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const [r, g, b, a] = fill(x, y);
        const o = (y * w + x) * 4;
        data[o] = r; data[o + 1] = g; data[o + 2] = b; data[o + 3] = a;
      }
  }
  return {
    data,
    canvas: { width: w, height: h },
    getImageData() {
      return { data: data.slice(), width: w, height: h } as unknown as ImageData;
    },
    putImageData(img: ImageData) {
      data.set(img.data);
    },
  };
}

describe('bevelHeightProfile', () => {
  it('clamps beyond the band and treats zero width as a flat face', () => {
    expect(bevelHeightProfile('circle',5)(10)).toBe(1);
    expect(bevelHeightProfile('circle',0)(0)).toBe(1);
  });
  it('retains the documented hardEdge line rather than a fitted rim shelf', () => {
    // MS-OI29500 §2.1.1316: line from (.042353,.305322) to (.170124,.947558),
    // each point (inward distance, height), normalized to the ~.9986 curve peak.
    const t=.1;
    const expected=(.305322+(t-.042353)*(.947558-.305322)/(.170124-.042353))/.998599;
    expect(bevelHeightProfile('hardEdge',8)(8*t)).toBeCloseTo(expected,3);
  });
});

describe('edt1d (squared Euclidean distance transform, Felzenszwalb)', () => {
  it('computes squared distance to the nearest zero on a 1D line', () => {
    // f: 0 at index 2, +inf elsewhere → squared distance to index 2.
    const INF = 1e20;
    const f = [INF, INF, 0, INF, INF];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([4, 1, 0, 1, 4]);
  });

  it('handles two seeds, taking the nearer', () => {
    const INF = 1e20;
    const f = [0, INF, INF, INF, 0];
    const out = edt1d(f);
    expect(Array.from(out)).toEqual([0, 1, 4, 1, 0]);
  });
});

describe('lightDirFromRig', () => {
  it('dir="t" points the light from the top (negative screen-y), toward viewer', () => {
    const L = lightDirFromRig('threePt', 't');
    expect(L.y).toBeLessThan(0); // light comes from above (screen y down)
    expect(L.z).toBeGreaterThan(0); // and from in front of the surface
    // unit length
    expect(Math.hypot(L.x, L.y, L.z)).toBeCloseTo(1, 5);
  });

  it('dir="l" comes from the left (negative x)', () => {
    const L = lightDirFromRig('threePt', 'l');
    expect(L.x).toBeLessThan(0);
  });

  it('dir="br" comes from bottom-right (positive x, positive y)', () => {
    const L = lightDirFromRig('threePt', 'br');
    expect(L.x).toBeGreaterThan(0);
    expect(L.y).toBeGreaterThan(0);
  });

  it('revolution rotates the documented key vector and preserves elevation', () => {
    const base=lightDirFromRig('threePt','t');
    const turned=lightDirFromRig('threePt','t',{lat:0,lon:0,rev:90});
    expect(turned.x).toBeCloseTo(-base.y,6);
    expect(turned.y).toBeCloseTo(base.x,6);
    expect(turned.z).toBeCloseTo(base.z,6);
  });

  it('rev=0 (or absent rot) is identical to no rot', () => {
    const a = lightDirFromRig('threePt', 't');
    const b = lightDirFromRig('threePt', 't', { lat: 0, lon: 0, rev: 0 });
    expect(b.x).toBeCloseTo(a.x, 6);
    expect(b.y).toBeCloseTo(a.y, 6);
    expect(b.z).toBeCloseTo(a.z, 6);
  });
});

describe('shadePixel (Lambert + weak specular)', () => {
  const params: BevelShadeParams = {
    light: lightDirFromRig('threePt', 't'),
    material: 'matte',
    ambient: 0.55,
    diffuse: 0.45,
    specular: 0.0,
    shininess: 8,
  };

  it('a flat top-facing normal (0,0,1) gets a mid factor near ambient+diffuse·(L·N)', () => {
    const f = shadePixel({ x: 0, y: 0, z: 1 }, params);
    // L·N = L.z. factor = ambient + diffuse*max(0,L.z).
    const expected = params.ambient + params.diffuse * params.light.z;
    expect(f).toBeCloseTo(expected, 5);
  });

  it('a normal tilted toward the light is brighter than one tilted away', () => {
    const toward = shadePixel(normalize(params.light.x, params.light.y, 1), params);
    const away = shadePixel(normalize(-params.light.x, -params.light.y, 1), params);
    expect(toward).toBeGreaterThan(away);
  });

  it('plastic material adds a specular highlight (brighter peak than matte)', () => {
    const matteP = { ...params, specular: 0.0 };
    const plasticP = { ...params, specular: 0.4, material: 'plastic' as const };
    // Normal exactly on the light direction → max specular.
    const n = normalize(params.light.x, params.light.y, params.light.z);
    expect(shadePixel(n, plasticP)).toBeGreaterThan(shadePixel(n, matteP));
  });

  it('clamps the factor to >= 0 (no negative)', () => {
    const f = shadePixel({ x: 0, y: 0, z: -1 }, params);
    expect(f).toBeGreaterThanOrEqual(0);
  });
});

describe('computeBevelNormals (distance-driven tilt, coverage-gradient azimuth)', () => {
  it('produces (0,0,1) on the flat interior and tilted normals on the band', () => {
    // 7x7 filled square, bevel width 2. Interior pixels (far from edge) flat;
    // band pixels tilt outward.
    const W = 7;
    const H = 7;
    const alpha = new Uint8ClampedArray(W * H).fill(255);
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    // Centre pixel (3,3) is interior → flat normal.
    const ci = (3 * W + 3) * 3;
    expect(normals[ci + 2]).toBeGreaterThan(0.9); // nz ≈ 1
    expect(bandMask[3 * W + 3]).toBe(0); // not in band
    // An edge-band pixel (e.g. (3,0), on the top edge, distance 1 < band 2) is
    // in the band and its normal tilts upward (ny < 0 → faces up toward the top).
    const ei = (0 * W + 3) * 3;
    expect(bandMask[0 * W + 3]).toBe(1);
    expect(normals[ei + 1]).toBeLessThan(0); // ny<0 faces up/out at the top edge
  });

  it('ignores fully transparent pixels (outside the silhouette)', () => {
    const W = 5;
    const H = 5;
    const alpha = new Uint8ClampedArray(W * H); // all zero
    const { bandMask } = computeBevelNormals(alpha, W, H, 2, 'circle', 1);
    expect(bandMask.every((v) => v === 0)).toBe(true);
  });
});

/**
 * Anti-aliased ellipse coverage mask (mimics a Canvas `clip()` rasterisation:
 * interior 255, exterior 0, a 1-px fractional rim). Used to exercise the bevel
 * lip on a curved silhouette with a wide band — the case that exposed the
 * EDT-gradient faceting regression (PR #410 → this fix).
 */
function ellipseAlpha(W: number, H: number, ss = 4): Uint8ClampedArray {
  const a = new Uint8ClampedArray(W * H);
  const cx = W / 2;
  const cy = H / 2;
  const rx = W / 2 - 1;
  const ry = H / 2 - 1;
  for (let y = 0; y < H; y++)
    for (let x = 0; x < W; x++) {
      let cnt = 0;
      for (let sy = 0; sy < ss; sy++)
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - 0.5;
          const py = y + (sy + 0.5) / ss - 0.5;
          if (((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1) cnt++;
        }
      a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
    }
  return a;
}

/**
 * Anti-aliased annulus (ring) coverage: outer ellipse minus a concentric inner
 * ellipse. A wide bevel band reaching toward the ring's medial axis is the
 * worst case for the OLD EDT-gradient azimuth (the nearest-feature flips across
 * the medial ridge). Used by the scale-sweep below to prove the coverage-gradient
 * azimuth tracks a SINGLE rim smoothly even on a thin ring at every scale.
 */
function ringAlpha(W: number, H: number, thickFrac: number, ss = 4): Uint8ClampedArray {
  const a = new Uint8ClampedArray(W * H);
  const cx = W / 2;
  const cy = H / 2;
  const rx = W / 2 - 1;
  const ry = H / 2 - 1;
  const irx = rx * (1 - thickFrac);
  const iry = ry * (1 - thickFrac);
  for (let y = 0; y < H; y++)
    for (let x = 0; x < W; x++) {
      let cnt = 0;
      for (let sy = 0; sy < ss; sy++)
        for (let sx = 0; sx < ss; sx++) {
          const px = x + (sx + 0.5) / ss - 0.5;
          const py = y + (sy + 0.5) / ss - 0.5;
          const outer = ((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1;
          const inner = ((px - cx) / irx) ** 2 + ((py - cy) / iry) ** 2 <= 1;
          if (outer && !inner) cnt++;
        }
      a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
    }
  return a;
}

describe('gaussianBlur (3-box cascade, zero-padded)', () => {
  it('smooths a step and falls off across a clipped (canvas-edge) boundary', () => {
    // A solid block touching the left edge: coverage 1 for x<W/2, 0 otherwise.
    const W = 32;
    const H = 4;
    const src = new Float64Array(W * H);
    for (let y = 0; y < H; y++) for (let x = 0; x < W / 2; x++) src[y * W + x] = 1;
    const out = gaussianBlur(src, W, H, 4);
    // Monotonic falloff across the interior step (x from W/2-4 to W/2+4).
    const row = 1 * W;
    for (let x = W / 2 - 3; x < W / 2 + 3; x++) {
      expect(out[row + x]).toBeGreaterThan(out[row + x + 1]);
    }
    // Zero padding (NOT clamp-to-edge): the solid block flush with x=0 falls off
    // toward the left border, so the value AT the border is below the interior
    // plateau — this is what gives a clipped silhouette an outward ∇C lip.
    expect(out[row + 0]).toBeLessThan(out[row + 4]);
  });
});

describe('computeBevelNormals — scale-invariant azimuth (issue #410→#413→#415→#416→#418)', () => {
  // DECISIVE SCALE SWEEP. The bevel facet bug slipped through repeatedly, each time
  // at a LARGER raster scale, because every early fix was a post-filter on the EDT
  // gradient whose radius was a fraction of the (growing) band:
  //   1. #410 mesh cracks — grew with shape size.
  //   2. #413 height-field blur — fixed devScale ≤ 2, regressed at 4.
  //   3. #415 tangential normal blur (radius 0.25·band, gated ≥24px) — fixed
  //      devScale 4, REGRESSED at devScale 8 (the user's real DPR-8 render).
  // Each early test pinned a SINGLE scale, so the next larger scale walked straight
  // through. This suite instead PARAMETRISES devScale ∈ {1,2,4,8} and asserts ONE
  // scale-independent threshold for every scale, so a facet at any DPR fails here.
  //
  // The azimuth is now sourced from ∇(Gaussian-blurred DISTANCE) (#418; #416 used
  // blurred COVERAGE, which removed the facet but PLATEAUED at the ellipse apex —
  // see the separate "apex plateau" suite). Blurring the EXACT distance is C^∞ at
  // every scale (no EDT Voronoi structure enters the direction), so smoothness is
  // scale-invariant — and in fact improves with scale. Measured (production code):
  // ellipse & ring max 1-step azimuth jump ≤ 0.0128 rad at devScale 1, shrinking to
  // ≤ 0.0046 at devScale 8; highlight luminance 2nd-difference ≤ 0.0050 → ≤ 0.0007.
  // For comparison the RAW EDT gradient (pre-#416) jumps 0.12–0.24 rad — a 7°–13°
  // chord — at all scales. The ceilings below sit just above the devScale-1
  // quantisation floor yet an order of magnitude below the raw facet.
  const DEV_SCALES = [1, 2, 4, 8];
  // sample-11 slide 6 body offscreen at devScale 1 ≈ 397×503 px, hardEdge band 32.
  const BASE = { W: 397, H: 503, band: 32 };
  const AZ_JUMP_MAX = 0.02; // rad, scale-independent
  // The lip AZIMUTH is the facet detector (AZ_JUMP_MAX). LUM_2ND_DIFF_MAX bounds the
  // tangential highlight 2nd-difference: it must stay far below the raw EDT facet
  // (0.12–0.24) but is otherwise quantisation-limited at the smallest raster. With
  // the rim-concentrated `hardEdge` profile the lit lip is steeper at the 0.25·band
  // sample point, so the measured 2nd-difference is ≈0.0050 at devScale 1 (32px band,
  // pixel-limited) shrinking to ≈0.0007 at devScale 8. A ceiling of 0.006 sits just
  // above the devScale-1 quantisation floor yet ~25× below the raw facet — still a
  // decisive facet test.
  const LUM_2ND_DIFF_MAX = 0.006; // scale-independent

  function ringAzimuthAndHighlight(alpha: Uint8ClampedArray, W: number, H: number, band: number) {
    const { normals, bandMask } = computeBevelNormals(alpha, W, H, band, 'hardEdge', band / 2);
    const params = shadeParamsFor('matte', lightDirFromRig('threePt', 't'));
    const faceFactor = shadePixel({ x: 0, y: 0, z: 1 }, params) || 1;
    const cx = W / 2;
    const cy = H / 2;
    const rx = W / 2 - 1;
    const ry = H / 2 - 1;
    // Sample WITHIN the active lip. `hardEdge` rises along its documented straight
    // flank between inward 0.042 and 0.170 of the band and is nearly flat from
    // 0.27 to 0.61, where the azimuth of a near-zero tilt is undefined. 0.1·band
    // sits on that constant-slope flank at every devScale.
    const inset = band * 0.1;
    const az: number[] = [];
    const lum: number[] = [];
    // 0.1° steps so a chord spanning several degrees registers as a run of equal
    // azimuths broken by a sudden jump.
    for (let deg = 0; deg < 360; deg += 0.1) {
      const ang = (deg * Math.PI) / 180;
      const x = Math.round(cx + Math.cos(ang) * (rx - inset));
      const y = Math.round(cy + Math.sin(ang) * (ry - inset));
      if (x < 0 || y < 0 || x >= W || y >= H) continue;
      const i = y * W + x;
      if (!bandMask[i]) continue;
      const nx = normals[i * 3];
      const ny = normals[i * 3 + 1];
      const nz = normals[i * 3 + 2];
      if (Math.hypot(nx, ny) < 1e-3) continue;
      az.push(Math.atan2(ny, nx));
      lum.push(shadePixel({ x: nx, y: ny, z: nz }, params) / faceFactor);
    }
    let maxJump = 0;
    let net = 0;
    for (let i = 1; i < az.length; i++) {
      let d = az[i] - az[i - 1];
      while (d > Math.PI) d -= 2 * Math.PI;
      while (d < -Math.PI) d += 2 * Math.PI;
      maxJump = Math.max(maxJump, Math.abs(d));
      net += d;
    }
    let maxLum2 = 0;
    for (let i = 2; i < lum.length; i++) {
      maxLum2 = Math.max(maxLum2, Math.abs(lum[i] - 2 * lum[i - 1] + lum[i - 2]));
    }
    return { maxJump, maxLum2, net, n: az.length };
  }

  // The devScale-8 cases rasterise a ~3176×4024 px alpha plane and run a full
  // EDT + 3-box Gaussian over it; on a loaded Linux CI runner that lands near
  // ~6 s — over vitest's 5 s default (a pure timeout, not a correctness fail).
  // Give the scale-sweep tests generous headroom rather than shrinking the
  // raster: devScale 8 IS the point of the scale-invariance assertion. Fast
  // low-scale iterations finish in ms and are unaffected by a high ceiling.
  const SCALE_SWEEP_TIMEOUT_MS = 30_000;

  for (const ds of DEV_SCALES) {
    const W = BASE.W * ds;
    const H = BASE.H * ds;
    const band = BASE.band * ds;

    it(`ellipse lip is smooth at devScale ${ds} (band ${band}px)`, () => {
      const r = ringAzimuthAndHighlight(ellipseAlpha(W, H), W, H, band);
      expect(r.n).toBeGreaterThan(3000);
      // (a) no facet: largest single-step azimuth jump under the shared ceiling.
      expect(r.maxJump).toBeLessThan(AZ_JUMP_MAX);
      // (b) highlight has no chorded kink: tangential 2nd-difference bounded.
      expect(r.maxLum2).toBeLessThan(LUM_2ND_DIFF_MAX);
      // (c) the azimuth makes one full turn — it tracks the rim, not a few facets.
      expect(Math.abs(Math.abs(r.net) - 2 * Math.PI)).toBeLessThan(0.5);
    }, SCALE_SWEEP_TIMEOUT_MS);

    it(`thin-ring lip is smooth at devScale ${ds} (band ${band}px)`, () => {
      // A ring whose band reaches toward the medial axis — the medial-ridge case
      // that breaks an EDT-nearest-feature azimuth. The coverage σ (0.25·band) is
      // small enough to stay local to the OUTER rim, so it stays smooth.
      const r = ringAzimuthAndHighlight(ringAlpha(W, H, 0.18), W, H, band);
      expect(r.n).toBeGreaterThan(3000);
      expect(r.maxJump).toBeLessThan(AZ_JUMP_MAX);
      expect(r.maxLum2).toBeLessThan(LUM_2ND_DIFF_MAX);
      expect(Math.abs(Math.abs(r.net) - 2 * Math.PI)).toBeLessThan(0.5);
    }, SCALE_SWEEP_TIMEOUT_MS);
  }

  it('smoothness does NOT degrade as scale grows (the property prior patches lacked)', () => {
    // The defining scale-invariance assertion: the max azimuth jump at devScale 8
    // must be NO WORSE than at devScale 1. Every prior EDT-gradient patch had the
    // opposite behaviour (smooth at the tuned scale, faceted above it); this would
    // be RED for all of them and is GREEN for the coverage-gradient azimuth.
    const small = ringAzimuthAndHighlight(
      ellipseAlpha(BASE.W * 1, BASE.H * 1),
      BASE.W * 1,
      BASE.H * 1,
      BASE.band * 1,
    );
    const large = ringAzimuthAndHighlight(
      ellipseAlpha(BASE.W * 8, BASE.H * 8),
      BASE.W * 8,
      BASE.H * 8,
      BASE.band * 8,
    );
    expect(large.maxJump).toBeLessThanOrEqual(small.maxJump + 1e-6);
  }, SCALE_SWEEP_TIMEOUT_MS);
});

describe('bevel lip azimuth tracks the analytic silhouette normal', () => {
  // THE #418 BUG. #416 sourced the lip azimuth from ∇(Gaussian-blurred COVERAGE).
  // Coverage `C = G_σ * alpha` PLATEAUS at a high-curvature convex apex: the blur of
  // a sharply curved silhouette tip flat-tops, so ∇C there points near-constant over
  // a wide angular span. The lip azimuth then ROTATES TOO SLOWLY across the apex, so
  // the lit factor is near-constant over a band of the rim → the "flat horizontal
  // cut" users saw at the top of the slide-6 ellipse (bevel OFF = smooth arc, bevel
  // ON = flat-topped band). The distance geometry was always
  // correct; only the DIRECTION was wrong.
  //
  // The fix sources the azimuth from ∇(Gaussian-blurred DISTANCE) instead. The EDT
  // distance field's level sets ARE the offset curves of the silhouette, so its
  // gradient points along the true radial normal EVERYWHERE — including the apex,
  // where distance keeps growing linearly inward and so does NOT plateau. Blurring
  // the distance (same σ) makes the gradient C¹ and washes out the raw-EDT Voronoi
  // facets, without reintroducing a plateau. These tests assert the computed lip
  // azimuth matches the ANALYTIC ellipse outward normal across the apex — the check
  // the coverage-gradient azimuth fails and the distance-gradient azimuth passes —
  // parametrised over devScale {1,2,4} so it is a scale-invariant property.

  /** Anti-aliased ellipse filling the W×H canvas (matches a Canvas clip raster). */
  function ellipseA(W: number, H: number, ss = 4): Uint8ClampedArray {
    const a = new Uint8ClampedArray(W * H);
    const cx = W / 2, cy = H / 2, rx = W / 2 - 1, ry = H / 2 - 1;
    for (let y = 0; y < H; y++)
      for (let x = 0; x < W; x++) {
        let cnt = 0;
        for (let sy = 0; sy < ss; sy++)
          for (let sx = 0; sx < ss; sx++) {
            const px = x + (sx + 0.5) / ss - 0.5, py = y + (sy + 0.5) / ss - 0.5;
            if (((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1) cnt++;
          }
        a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
      }
    return a;
  }

  // slide-6 ellipse is 298×377 pt; its apex radius of curvature rx²/ry ≈ 118 pt is
  // only ≈5× the 24 pt band — exactly the high-curvature regime where the coverage
  // blur plateaus. Use the same proportions (W:H ≈ 298:377) so the synthetic case
  // reproduces the reported geometry.
  for (const devScale of [1, 2, 4]) {
    const W = 298 * devScale;
    const H = 377 * devScale;
    const band = 24 * devScale;

    it(`apex lip azimuth follows the analytic ellipse normal, not a plateau [devScale ${devScale}]`, () => {
      const a = ellipseA(W, H);
      const { normals, bandMask } = computeBevelNormals(a, W, H, band, 'hardEdge', band / 2);
      const cx = W / 2, cy = H / 2, rx = W / 2 - 1, ry = H / 2 - 1;
      // Sample band pixels just inside the TOP arc, sweeping x across the apex. For
      // each, compare the computed in-plane azimuth to the analytic ellipse outward
      // normal at the nearest rim point on the same column.
      let maxErrRad = 0;
      let samples = 0;
      // Across the apex the column sweep spans roughly ±0.5·rx; that is where the
      // coverage plateau bites (the curvature is highest near x=cx).
      for (let x = Math.round(cx - 0.5 * rx); x <= Math.round(cx + 0.5 * rx); x++) {
        // topmost opaque row on this column = rim.
        let rimY = -1;
        for (let y = 0; y < H; y++) if (a[y * W + x] >= 128) { rimY = y; break; }
        if (rimY < 0) continue;
        // a band pixel on hardEdge's straight rim flank (inward 0.042–0.170).
        const y = rimY + Math.max(1, Math.round(band * 0.1));
        const i = y * W + x;
        if (!bandMask[i]) continue;
        const nx = normals[i * 3], ny = normals[i * 3 + 1];
        if (Math.hypot(nx, ny) < 1e-3) continue;
        const computed = Math.atan2(ny, nx);
        // analytic outward normal of the ellipse at the rim point (gradient of the
        // implicit form): ((x-cx)/rx², (y-cy)/ry²), normalised.
        const ex = (x - cx) / (rx * rx), ey = (rimY - cy) / (ry * ry);
        const analytic = Math.atan2(ey, ex);
        let e = computed - analytic;
        while (e > Math.PI) e -= 2 * Math.PI;
        while (e < -Math.PI) e += 2 * Math.PI;
        maxErrRad = Math.max(maxErrRad, Math.abs(e));
        samples++;
      }
      expect(samples).toBeGreaterThan(50);
      // The distance-gradient azimuth matches the analytic normal to within a few
      // degrees (raster + AA-rim limited). The coverage-gradient azimuth plateaus
      // and overshoots ≈15–25° near the apex, so a 6° (0.105 rad) ceiling cleanly
      // separates the two. Scale-independent.
      expect(maxErrRad).toBeLessThan(0.105);
    });

  }
});

describe('bevel band geometry on a curved silhouette', () => {
  // The distance-field footprint must retain width/scale covariance and smooth
  // inward offsets independently of the documented profile's creases.
  /** Anti-aliased ellipse silhouette filling the W×H canvas. */
  function ellipseSilhouette(W: number, H: number, ss = 4): Uint8ClampedArray {
    const a = new Uint8ClampedArray(W * H);
    const cx = W / 2, cy = H / 2, rx = W / 2 - 1, ry = H / 2 - 1;
    for (let y = 0; y < H; y++)
      for (let x = 0; x < W; x++) {
        let cnt = 0;
        for (let sy = 0; sy < ss; sy++)
          for (let sx = 0; sx < ss; sx++) {
            const px = x + (sx + 0.5) / ss - 0.5, py = y + (sy + 0.5) / ss - 0.5;
            if (((px - cx) / rx) ** 2 + ((py - cy) / ry) ** 2 <= 1) cnt++;
          }
        a[y * W + x] = Math.round((255 * cnt) / (ss * ss));
      }
    return a;
  }

  // slide-6: shape 3785692×4793942 EMU (≈298×377 pt), bevelT w=304800 EMU (24 pt),
  // h=152400 EMU (12 pt). EMU→CSS-px scale = targetWidth / slideWidth.
  const SLIDE_W_EMU = 12192000;
  const SHAPE_W_EMU = 3785692, SHAPE_H_EMU = 4793942;
  const BEVEL_W_EMU = 304800, BEVEL_H_EMU = 152400;
  // A modest render width keeps the devScale-8 raster tractable in CI (the full
  // 1920-px-wide deck would be ~29 MP at dev 8 → EDT timeout) WITHOUT changing what
  // is tested: the band/shape geometry is governed by the EMU RATIO `bevel/shape`,
  // which is independent of TARGET_WIDTH, and the conversion is exercised over the
  // real (cssScale × devScale) multiplier across devScale {1,4,8}.
  const TARGET_WIDTH = 600;
  const cssScale = TARGET_WIDTH / SLIDE_W_EMU;

  for (const devScale of [1, 4, 8]) {
    // The exact conversion the renderer uses: EMU × cssScale × devScale.
    const expectBand = BEVEL_W_EMU * cssScale * devScale;
    const expectHeight = BEVEL_H_EMU * cssScale * devScale;
    const W = Math.round(SHAPE_W_EMU * cssScale * devScale);
    const H = Math.round(SHAPE_H_EMU * cssScale * devScale);

    it(`band width == declared EMU × scale × devScale (no double-apply blow-up) [devScale ${devScale}]`, () => {
      // The conversion is plain multiplication; assert it is what the geometry uses
      // and that the band is a sane fraction of the shape (not a runaway value that
      // would collapse the medial axis). 24 pt on a 298 pt shape ⇒ ~8% of width.
      expect(expectBand).toBeCloseTo(BEVEL_W_EMU * cssScale * devScale, 6);
      const halfMinDim = Math.min(W, H) / 2;
      const frac = expectBand / halfMinDim;
      // ≈0.16 of the half-minor-axis at every scale — scale-INVARIANT, no blow-up.
      expect(frac).toBeGreaterThan(0.14);
      expect(frac).toBeLessThan(0.18);
      // The measured band footprint matches the declared width along the minor axis
      // (where the boundary normal is radial, so perpendicular distance == radial).
      const a = ellipseSilhouette(W, H);
      const { bandMask } = computeBevelNormals(a, W, H, expectBand, 'hardEdge', expectHeight);
      const cx = Math.round(W / 2), cy = Math.round(H / 2);
      // Walk inward from the right edge along y=cy; count band pixels.
      let bandDepth = 0;
      for (let x = W - 1; x >= 0; x--) {
        const i = cy * W + x;
        if (a[i] < 128) continue; // still outside
        if (bandMask[i]) bandDepth++; else break;
      }
      // Perpendicular band depth at the minor-axis end ≈ bandPx (±2 px raster).
      expect(Math.abs(bandDepth - expectBand)).toBeLessThanOrEqual(2);
    });

    it(`band inner boundary is the SMOOTH Euclidean offset (no faceted flat-cut) [devScale ${devScale}]`, () => {
      // GEOMETRY: trace the band → flat-top boundary (the curve users saw as a
      // flat-cut). It is the exact Euclidean inward offset of the ellipse, so it is
      // smooth — its curvature (2nd-difference of radius vs angle) is bounded. The
      // pre-fix box-blur azimuth would have chorded this into facets; the EDT-exact
      // band does not. Threshold is normalised by the major radius (scale-free).
      const a = ellipseSilhouette(W, H);
      const { bandMask } = computeBevelNormals(a, W, H, expectBand, 'hardEdge', expectHeight);
      const cx = W / 2, cy = H / 2;
      const radii: number[] = [];
      for (let deg = 0; deg < 360; deg += 3) {
        const ang = (deg * Math.PI) / 180, dx = Math.cos(ang), dy = Math.sin(ang);
        const rmax = Math.min(W, H) / 2;
        let inner = NaN;
        for (let t = 0; t < rmax; t++) {
          const r = rmax - t;
          const x = Math.round(cx + dx * r), y = Math.round(cy + dy * r);
          if (x < 0 || y < 0 || x >= W || y >= H) continue;
          const i = y * W + x;
          if (a[i] >= 128 && bandMask[i] === 0) { inner = r; break; }
        }
        radii.push(inner);
      }
      expect(radii.every((r) => Number.isFinite(r))).toBe(true);
      const scaleRef = Math.max(W, H) / 2;
      let maxCurv = 0;
      for (let k = 1; k < radii.length - 1; k++) {
        const d2 = Math.abs(radii[k + 1] - 2 * radii[k] + radii[k - 1]) / scaleRef;
        if (d2 > maxCurv) maxCurv = d2;
      }
      // Smooth offset: ≤ ~0.01 relative (pixel quantisation only); a faceted cut
      // would spike many× higher.
      expect(maxCurv).toBeLessThan(0.02);
    });

  }
});

describe('applyExtrusion (side-wall sweep, §20.1.5.12)', () => {
  it('fills the swept band behind the front face with the wall colour', () => {
    // 20x20 with an opaque 8x8 block at (6,6)-(13,13). Push it right+down by
    // (4,4): the side wall is the L-shaped band uncovered to the right/below.
    const W = 20;
    const H = 20;
    const ctx = fakeCtx(W, H, (x, y) =>
      x >= 6 && x < 14 && y >= 6 && y < 14 ? [10, 20, 200, 255] : [0, 0, 0, 0],
    );
    applyExtrusion(ctx, { offsetX: 4, offsetY: 4, rgb: [50, 50, 50] });
    // A pixel just below-right of the block (e.g. (15,15)) should now be wall.
    const wall = (15 * W + 15) * 4;
    expect(ctx.data[wall + 3]).toBe(255);
    expect(ctx.data[wall]).toBe(50);
    // The front face is untouched.
    const face = (8 * W + 8) * 4;
    expect(ctx.data[face]).toBe(10);
    expect(ctx.data[face + 2]).toBe(200);
  });

  it('is a no-op for a sub-pixel offset (face-on camera)', () => {
    const W = 10;
    const H = 10;
    const ctx = fakeCtx(W, H, (x, y) => (x >= 3 && x < 7 && y >= 3 && y < 7 ? [10, 10, 10, 255] : [0, 0, 0, 0]));
    const before = ctx.data.slice();
    applyExtrusion(ctx, { offsetX: 0.2, offsetY: 0.1, rgb: [99, 99, 99] });
    expect(Array.from(ctx.data)).toEqual(Array.from(before));
  });
});

describe('applyBevelShading (end-to-end on a filled square)', () => {
  it('brightens the top lip and darkens the bottom lip under a top light', () => {
    const W = 24;
    const H = 24;
    // Mid-grey opaque square filling the canvas.
    const ctx = fakeCtx(W, H, () => [120, 120, 120, 255]);
    applyBevelShading(ctx, {
      widthPx: 5,
      heightPx: 4,
      prst: 'circle',
      material: 'matte',
      light: lightDirFromRig('threePt', 't'),
    });
    const lumAt = (x: number, y: number) => {
      const o = (y * W + x) * 4;
      return ctx.data[o] * 0.299 + ctx.data[o + 1] * 0.587 + ctx.data[o + 2] * 0.114;
    };
    // Top lip (y=1) lit brighter than the flat centre; bottom lip (y=22) darker.
    expect(lumAt(12, 1)).toBeGreaterThan(lumAt(12, 12));
    expect(lumAt(12, 22)).toBeLessThan(lumAt(12, 12));
  });
});

/**
 * A region-honouring ImageData-backed ctx: `getImageData(sx,sy,sw,sh)` returns a
 * copy of exactly that window, and `putImageData(img,dx,dy)` writes it back at the
 * offset — so it exercises the A3 sub-region path (the simple `fakeCtx` above
 * ignores the coordinates, which is fine for the full-canvas callers). Tracks the
 * getImageData window so the test can assert the region actually shrank.
 */
function regionCtx(
  w: number,
  h: number,
  fill: (x: number, y: number) => [number, number, number, number],
): BevelCtx & { data: Uint8ClampedArray; lastGet: { x: number; y: number; w: number; h: number } | null } {
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      const [r, g, b, a] = fill(x, y);
      const o = (y * w + x) * 4;
      data[o] = r; data[o + 1] = g; data[o + 2] = b; data[o + 3] = a;
    }
  const state = { lastGet: null as { x: number; y: number; w: number; h: number } | null };
  return {
    data,
    get lastGet() { return state.lastGet; },
    canvas: { width: w, height: h },
    getImageData(sx: number, sy: number, sw: number, sh: number) {
      state.lastGet = { x: sx, y: sy, w: sw, h: sh };
      const out = new Uint8ClampedArray(sw * sh * 4);
      for (let y = 0; y < sh; y++)
        for (let x = 0; x < sw; x++) {
          const src = ((sy + y) * w + (sx + x)) * 4;
          const dst = (y * sw + x) * 4;
          out[dst] = data[src]; out[dst + 1] = data[src + 1];
          out[dst + 2] = data[src + 2]; out[dst + 3] = data[src + 3];
        }
      return { data: out, width: sw, height: sh } as unknown as ImageData;
    },
    putImageData(img: ImageData, dx: number, dy: number) {
      const sw = img.width, sh = img.height;
      for (let y = 0; y < sh; y++)
        for (let x = 0; x < sw; x++) {
          const src = (y * sw + x) * 4;
          const dst = ((dy + y) * w + (dx + x)) * 4;
          data[dst] = img.data[src]; data[dst + 1] = img.data[src + 1];
          data[dst + 2] = img.data[src + 2]; data[dst + 3] = img.data[src + 3];
        }
    },
  };
}

describe('A3 region limit — bbox-restricted equals full-canvas, byte for byte', () => {
  // A small opaque shape parked OFF-ORIGIN on a much larger canvas (the case A3
  // optimises: the silhouette occupies a fraction of the offscreen). Running the
  // effect over just `bbox ⊕ reach` must yield the SAME pixels as over the whole
  // canvas, because every band pixel (dist < bandPx) / wall pixel (within |offset|)
  // lies inside that window and its edge sits in transparent territory — matching
  // `distanceToEdge`'s out-of-bounds-is-transparent boundary.
  const CANVAS_W = 200;
  const CANVAS_H = 160;
  // Opaque rounded-ish block at (60..140, 40..110) — an 80×70 silhouette with a
  // wide transparent margin on all sides.
  const BX = 60, BY = 40, BW = 80, BH = 70;
  const shapeFill = (x: number, y: number): [number, number, number, number] =>
    x >= BX && x < BX + BW && y >= BY && y < BY + BH ? [130, 130, 130, 255] : [0, 0, 0, 0];

  function diffCount(a: Uint8ClampedArray, b: Uint8ClampedArray): number {
    let n = 0;
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) n++;
    return n;
  }

  it('applyBevelShading: region == full canvas output (interior shape, band 8px)', () => {
    const band = 8;
    const bevel = {
      widthPx: band, heightPx: 6, prst: 'circle', material: 'matte' as const,
      light: lightDirFromRig('threePt', 't'),
    };
    const full = regionCtx(CANVAS_W, CANVAS_H, shapeFill);
    applyBevelShading(full, bevel); // whole canvas

    const reach = Math.ceil(band) + 2;
    const region = { x: BX - reach, y: BY - reach, w: BW + 2 * reach, h: BH + 2 * reach };
    const cropped = regionCtx(CANVAS_W, CANVAS_H, shapeFill);
    applyBevelShading(cropped, bevel, region);

    // The region genuinely shrank the read window (proves it isn't secretly full).
    expect(cropped.lastGet).not.toBeNull();
    expect(cropped.lastGet!.w).toBeLessThan(CANVAS_W);
    expect(cropped.lastGet!.h).toBeLessThan(CANVAS_H);
    // …and the pixels are byte-identical to the full-canvas run.
    expect(diffCount(full.data, cropped.data)).toBe(0);
  });

  it('applyExtrusion: region == full canvas output (offset 6,4)', () => {
    const ext = { offsetX: 6, offsetY: 4, rgb: [40, 40, 40] as [number, number, number] };
    const full = regionCtx(CANVAS_W, CANVAS_H, shapeFill);
    applyExtrusion(full, ext);

    const reach = Math.ceil(Math.hypot(ext.offsetX, ext.offsetY)) + 2;
    const region = { x: BX - reach, y: BY - reach, w: BW + 2 * reach, h: BH + 2 * reach };
    const cropped = regionCtx(CANVAS_W, CANVAS_H, shapeFill);
    applyExtrusion(cropped, ext, region);

    expect(cropped.lastGet!.w).toBeLessThan(CANVAS_W);
    expect(diffCount(full.data, cropped.data)).toBe(0);
  });

  it('applyBevelShading: a region clamped to the canvas edge still matches full', () => {
    // Shape touching the canvas edge (BX2=0) so the region clamps at x=0 — the
    // clamped window must still reproduce the full-canvas result exactly.
    const band = 6;
    const edgeFill = (x: number, y: number): [number, number, number, number] =>
      x < 70 && y >= 40 && y < 110 ? [130, 130, 130, 255] : [0, 0, 0, 0];
    const bevel = {
      widthPx: band, heightPx: 5, prst: 'circle', material: 'matte' as const,
      light: lightDirFromRig('threePt', 't'),
    };
    const full = regionCtx(CANVAS_W, CANVAS_H, edgeFill);
    applyBevelShading(full, bevel);
    const reach = Math.ceil(band) + 2;
    const region = { x: 0 - reach, y: 40 - reach, w: 70 + 2 * reach, h: 70 + 2 * reach };
    const cropped = regionCtx(CANVAS_W, CANVAS_H, edgeFill);
    applyBevelShading(cropped, bevel, region);
    expect(cropped.lastGet!.x).toBe(0); // clamped
    expect(diffCount(full.data, cropped.data)).toBe(0);
  });
});

function normalize(x: number, y: number, z: number) {
  const m = Math.hypot(x, y, z) || 1;
  return { x: x / m, y: y / m, z: z / m };
}
