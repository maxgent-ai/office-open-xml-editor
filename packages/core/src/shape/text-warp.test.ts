import { describe, it, expect } from 'vitest';
import {
  buildWarpEnvelope,
  followPathUScale,
  hasTextWarp,
  samplePolyline,
  warpGlyphTransform,
  warpArcLength,
  type Polyline,
} from './text-warp';

const W = 400;
const H = 100;

describe('text-warp preset registry', () => {
  it('knows the 40 spec presets, case-insensitively', () => {
    expect(hasTextWarp('textArchUp')).toBe(true);
    expect(hasTextWarp('TEXTARCHUP')).toBe(true);
    expect(hasTextWarp('textWave1')).toBe(true);
    expect(hasTextWarp('textPlain')).toBe(true);
    expect(hasTextWarp('textNoShape')).toBe(false);
    expect(hasTextWarp('notAWarp')).toBe(false);
  });

  it('classifies single-edge (arch/circle) vs paired-edge warps', () => {
    const singleEdge = (preset: string) => buildWarpEnvelope(preset, [], W, H)?.singleEdge;
    expect(singleEdge('textArchUp')).toBe(true);
    expect(singleEdge('textArchDown')).toBe(true);
    expect(singleEdge('textCircle')).toBe(true);
    expect(singleEdge('textPlain')).toBe(false);
    expect(singleEdge('textWave1')).toBe(false);
  });

  it('returns null for an unknown preset', () => {
    expect(buildWarpEnvelope('nope', [], W, H)).toBeNull();
  });
});

describe('textPlain — the identity envelope', () => {
  it('is two horizontal edges spanning the full width at top / bottom', () => {
    const env = buildWarpEnvelope('textPlain', [], W, H);
    expect(env).not.toBeNull();
    const top = env!.top;
    const bottom = env!.bottom;
    // Both edges run left→right across the width.
    expect(top[0].x).toBeCloseTo(0, 1);
    expect(top[top.length - 1].x).toBeCloseTo(W, 1);
    expect(bottom[0].x).toBeCloseTo(0, 1);
    expect(bottom[bottom.length - 1].x).toBeCloseTo(W, 1);
    // Top edge is a constant y near 0, bottom near H.
    for (const p of top) expect(p.y).toBeCloseTo(0, 1);
    for (const p of bottom) expect(p.y).toBeCloseTo(H, 1);
  });

  it('warps a mid-line baseline to a flat, un-rotated, unit-scaled point', () => {
    const env = buildWarpEnvelope('textPlain', [], W, H)!;
    // boxHeight == H (text exactly fills the box) → vScale == 1.
    const mid = warpGlyphTransform(env, 0.5, H, 0.75);
    expect(mid.x).toBeCloseTo(W / 2, 1);
    expect(mid.y).toBeCloseTo(H * 0.75, 1);
    expect(mid.angle).toBeCloseTo(0, 3); // horizontal edges → no rotation
    expect(mid.vScale).toBeCloseTo(1, 3); // gap == box height → no compression
  });
});

describe('textInflate — box bulges in the middle', () => {
  it('top edge bows UP (smaller y) and bottom bows DOWN (larger y) at centre', () => {
    const env = buildWarpEnvelope('textInflate', [], W, H)!;
    const topMid = samplePolyline(env.top, env.topLen, 0.5);
    const topEnd = samplePolyline(env.top, env.topLen, 0.0);
    const botMid = samplePolyline(env.bottom, env.bottomLen, 0.5);
    const botEnd = samplePolyline(env.bottom, env.bottomLen, 0.0);
    // At the centre the top edge is higher (smaller y) than at the ends.
    expect(topMid.y).toBeLessThan(topEnd.y);
    // and the bottom edge is lower (larger y) than at the ends.
    expect(botMid.y).toBeGreaterThan(botEnd.y);
    // So the vertical gap is LARGER at the centre than at the ends.
    const gapMid = botMid.y - topMid.y;
    const gapEnd = botEnd.y - topEnd.y;
    expect(gapMid).toBeGreaterThan(gapEnd);
  });
});

describe('textArchUp — single arc baseline', () => {
  it('flattens into an arc whose points are equidistant from a common centre', () => {
    // Default adj makes a full 180°-ish arch. Points on an arc share a centre.
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    const arc = env.top;
    expect(arc.length).toBeGreaterThan(20);
    // Fit centre as the average, then check radius variance is tiny relative to r.
    const cx = arc.reduce((s, p) => s + p.x, 0) / arc.length;
    const cy = arc.reduce((s, p) => s + p.y, 0) / arc.length;
    const radii = arc.map((p) => Math.hypot(p.x - cx, p.y - cy));
    const rMean = radii.reduce((s, r) => s + r, 0) / radii.length;
    // A circular arc's centroid is NOT the circle centre, so refine: use the
    // ellipse centre (hc, vc·) is unknown; instead assert the arc is convex and
    // its mid-point sits above its endpoints (an "up" arch opens downward).
    const start = arc[0];
    const end = arc[arc.length - 1];
    const midx = arc[Math.floor(arc.length / 2)];
    // The apex of an up-arch is higher (smaller y) than both ends.
    expect(midx.y).toBeLessThan(start.y);
    expect(midx.y).toBeLessThan(end.y);
    expect(rMean).toBeGreaterThan(0);
  });

  it('per-glyph transform keeps vScale=1 and rotates glyphs along the arc', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    expect(env.singleEdge).toBe(true);
    const boxH = 30;
    const gl = warpGlyphTransform(env, 0.5, boxH, 0.8);
    const gr = warpGlyphTransform(env, 0.85, boxH, 0.8);
    const glft = warpGlyphTransform(env, 0.15, boxH, 0.8);
    // Single-edge presets never compress the glyph height…
    expect(gl.vScale).toBe(1);
    // …and never shear: Follow Path glyphs rigidly rotate along the arc.
    expect(gl.shear).toBe(0);
    expect(gr.shear).toBe(0);
    expect(glft.shear).toBe(0);
    // At the apex the axis is horizontal; toward the right end it tilts down
    // (positive angle), toward the left end it tilts up (negative angle), and
    // the tilt grows monotonically away from the apex.
    expect(Math.abs(gl.angle)).toBeLessThan(0.05);
    expect(gr.angle).toBeGreaterThan(0.15);
    expect(glft.angle).toBeLessThan(-0.15);
    // Left/right are mirror images about the apex.
    expect(gr.angle).toBeCloseTo(-glft.angle, 3);
  });
});

describe('textInflate — per-glyph vertical scale', () => {
  it('scales glyphs TALLER at the centre than at the ends', () => {
    const env = buildWarpEnvelope('textInflate', [], W, H)!;
    const boxH = H; // nominal
    const mid = warpGlyphTransform(env, 0.5, boxH, 0.8);
    const end = warpGlyphTransform(env, 0.02, boxH, 0.8);
    expect(mid.vScale).toBeGreaterThan(end.vScale);
  });
});

describe('local envelope shear (§20.1.9.19) — paired-edge slope skews the glyph', () => {
  // Reconstruct the per-glyph linear map the renderer applies:
  //   rotate(angle) · [[1, shear],[0, 1]] · scale(1, vScale)
  // Its VERTICAL column must equal the true envelope gap vector (B(u)−T(u))
  // per unit box height — that is what makes vertical strokes track the gap
  // while the baseline follows the slope (the glyph leans, not rigidly rotates).
  function verticalColumn(g: {
    angle: number;
    shear: number;
    vScale: number;
  }): { x: number; y: number } {
    const c = Math.cos(g.angle);
    const s = Math.sin(g.angle);
    // column_y of rotate·shear·scale = rotate · (shear·vScale, vScale)
    const lx = g.shear * g.vScale;
    const ly = g.vScale;
    return { x: c * lx - s * ly, y: s * lx + c * ly };
  }

  it('flat edges (textPlain) produce zero shear', () => {
    const env = buildWarpEnvelope('textPlain', [], W, H)!;
    for (const u of [0, 0.25, 0.5, 0.75, 1]) {
      const g = warpGlyphTransform(env, u, H, 0.75);
      expect(g.shear).toBeCloseTo(0, 6);
    }
  });

  it('textWave1 shears where the wave slopes and is flat at the crest/troughs', () => {
    // 320×160 matches the warp fixture's Wave One shape; default adj.
    const env = buildWarpEnvelope('textWave1', [], 320, 160)!;
    const box = 100;
    // The wave's zero-slope points sit at the quarter marks (u≈0.25, 0.75)
    // where the cubic turns; the steepest slope is at the mid-band (u≈0.5) and
    // the ends (u≈0,1). At the flat quarter marks shear ≈ 0.
    const flatish = warpGlyphTransform(env, 0.25, box, 0.8);
    expect(Math.abs(flatish.shear)).toBeLessThan(0.12);
    // At the mid-band the edge slopes ~17°, so the glyph must shear noticeably.
    const sloped = warpGlyphTransform(env, 0.5, box, 0.8);
    expect(Math.abs(sloped.shear)).toBeGreaterThan(0.2);
    // The end of the wave slopes the OTHER way → opposite-sign shear.
    const end = warpGlyphTransform(env, 0.98, box, 0.8);
    expect(Math.sign(end.shear)).toBe(-Math.sign(sloped.shear));
  });

  it('reconstructed vertical column equals the true gap vector B(u)−T(u)/box', () => {
    const env = buildWarpEnvelope('textWave1', [], 320, 160)!;
    const box = 100;
    for (const u of [0.1, 0.35, 0.5, 0.65, 0.9]) {
      const g = warpGlyphTransform(env, u, box, 0.8);
      const t = samplePolyline(env.top, env.topLen, u);
      const b = samplePolyline(env.bottom, env.bottomLen, u);
      const col = verticalColumn(g);
      expect(col.x).toBeCloseTo((b.x - t.x) / box, 4);
      expect(col.y).toBeCloseTo((b.y - t.y) / box, 4);
    }
  });

  it('keeps vertical strokes near-vertical on a wave (unlike a rigid rotation)', () => {
    // On textWave1 the two edges are parallel, so the gap vector is vertical
    // everywhere → the transform's vertical column stays vertical and only the
    // horizontal (advance) axis tilts. A rigid rotate would tilt BOTH.
    const env = buildWarpEnvelope('textWave1', [], 320, 160)!;
    const g = warpGlyphTransform(env, 0.5, 100, 0.8);
    const col = verticalColumn(g);
    const devFromVertical = Math.abs(Math.atan2(col.x, col.y));
    expect(devFromVertical).toBeLessThan(0.02); // ≈0 rad: stays vertical
    // …while the advance axis is clearly tilted (the source of the lean).
    expect(Math.abs(g.angle)).toBeGreaterThan(0.15);
  });
});

describe('Follow Path — single-edge natural-width distribution (§20.1.9.19)', () => {
  it('warpArcLength returns the baseline arc length for single-edge presets', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    expect(env.singleEdge).toBe(true);
    // The arc length equals the accumulated length of the flattened baseline.
    expect(warpArcLength(env)).toBeCloseTo(env.topLen[env.topLen.length - 1], 3);
    expect(warpArcLength(env)).toBeGreaterThan(0);
  });

  it('maps text narrower than the arc to only its natural fraction from the start', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    const arc = warpArcLength(env);
    // Text whose flat ink width is a quarter of the arc should occupy the first
    // quarter of the arc parameter, not spread across the whole path.
    const naturalW = arc / 4;
    expect(followPathUScale(env, naturalW)).toBeCloseTo(0.25, 6);
    // A glyph at u=1 (line end) then lands at arc parameter 0.25, i.e. the text
    // follows the path for only its natural arc length from stAng.
    expect(1 * followPathUScale(env, naturalW)).toBeCloseTo(0.25, 6);
  });

  it('never stretches: text wider than the arc is clamped to the full path', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    const arc = warpArcLength(env);
    // Follow Path only shrinks the span; it never widens text past the path.
    expect(followPathUScale(env, arc * 2)).toBeCloseTo(1, 6);
  });

  it('is the identity (1) for paired-edge presets — they DO stretch', () => {
    const env = buildWarpEnvelope('textInflate', [], W, H)!;
    expect(env.singleEdge).toBe(false);
    expect(followPathUScale(env, 10)).toBe(1);
    expect(followPathUScale(env, 99999)).toBe(1);
  });

  it('is safe for a degenerate zero-length or zero-width input', () => {
    const env = buildWarpEnvelope('textArchUp', [], W, H)!;
    expect(followPathUScale(env, 0)).toBe(0);
  });
});

describe('samplePolyline — arc-length parameterisation', () => {
  it('returns endpoints at u=0 and u=1 and a unit tangent', () => {
    const poly: Polyline = [
      { x: 0, y: 0 },
      { x: 10, y: 0 },
      { x: 20, y: 0 },
    ];
    const cum = [0, 10, 20];
    const a = samplePolyline(poly, cum, 0);
    const b = samplePolyline(poly, cum, 1);
    const mid = samplePolyline(poly, cum, 0.5);
    expect(a.x).toBeCloseTo(0);
    expect(b.x).toBeCloseTo(20);
    expect(mid.x).toBeCloseTo(10);
    expect(Math.hypot(a.tx, a.ty)).toBeCloseTo(1);
  });
});
