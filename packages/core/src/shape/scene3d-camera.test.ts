import { describe, it, expect } from 'vitest';
import {
  computeScene3dQuad,
  isScene3dNonIdentity,
  computeDepthOffset,
  type CameraInput,
  type Vec2,
} from './scene3d-camera';

const W = 400;
const H = 300;

/** Round a Vec2 for readable snapshots / approximate comparisons. */
function r(p: Vec2, d = 3): { x: number; y: number } {
  const k = 10 ** d;
  return { x: Math.round(p.x * k) / k, y: Math.round(p.y * k) / k };
}

describe('computeScene3dQuad', () => {
  it('orthographicFront with no rot is the identity rectangle', () => {
    const cam: CameraInput = { prst: 'orthographicFront' };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isIdentity).toBe(true);
    expect(q.isAffine).toBe(true);
    expect(r(q.corners[0])).toEqual({ x: 0, y: 0 });
    expect(r(q.corners[1])).toEqual({ x: W, y: 0 });
    expect(r(q.corners[2])).toEqual({ x: W, y: H });
    expect(r(q.corners[3])).toEqual({ x: 0, y: H });
  });

  it('orthographicFront with all-zero rot is still identity', () => {
    const cam: CameraInput = { prst: 'orthographicFront', rot: { lat: 0, lon: 0, rev: 0 } };
    expect(computeScene3dQuad(cam, W, H).isIdentity).toBe(true);
    expect(isScene3dNonIdentity(cam)).toBe(false);
  });

  it('positions the back surface at the extrusion base rather than at the front plane', () => {
    const camera = { prst: 'orthographicFront', rot: { lat: 0, lon: 150, rev: 0 } };
    const front = computeScene3dQuad(camera, 100, 200);
    const back = computeScene3dQuad(camera, 100, 200, 40);
    // A 40-unit push along -Z rotated 150 degrees moves every corner 20
    // units right in parallel projection: 40 * sin(150 degrees).
    for (let i = 0; i < 4; i++) {
      expect(back.corners[i].x - front.corners[i].x).toBeCloseTo(20, 10);
      expect(back.corners[i].y).toBeCloseTo(front.corners[i].y, 10);
    }
  });

  it('unknown preset falls back to identity (no throw)', () => {
    const cam: CameraInput = { prst: 'unknownCamera' };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isIdentity).toBe(true);
  });

  it('recognizes every named perspective preset as perspective projection', () => {
    const cam: CameraInput = {
      prst: 'perspectiveContrastingLeftFacing',
      rot: { lat: 30, lon: 40, rev: 350 },
    };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
  });

  it('lon>0 turns the right edge TOWARD the viewer (right edge nearer/taller)', () => {
    // A pure longitude (Y-axis) turn. With lon > 0 the right edge comes toward
    // the viewer and the left edge recedes (file-angle convention — see the
    // module header). The left and right edges stay vertical (equal x for the
    // two corners on each side) and the quad mirrors top↔bottom about the
    // horizontal centre line; the near (right) edge is TALLER than the far (left)
    // edge by perspective foreshortening.
    const cam: CameraInput = { prst: 'perspectiveFront', rot: { lat: 0, lon: 25, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
    const [tl, tr, br, bl] = q.corners.map((c) => r(c));
    // Left edge vertical (TL.x == BL.x); right edge vertical (TR.x == BR.x).
    expect(tl.x).toBeCloseTo(bl.x, 2);
    expect(tr.x).toBeCloseTo(br.x, 2);
    // Mirror about the horizontal centre line: TL.y↔BL.y, TR.y↔BR.y.
    expect(tl.y).toBeCloseTo(H - bl.y, 2);
    expect(tr.y).toBeCloseTo(H - br.y, 2);
    // Right (near) edge taller than left (far) edge.
    const leftH = bl.y - tl.y;
    const rightH = br.y - tr.y;
    expect(rightH).toBeGreaterThan(leftH + 1);
  });

  it('lat>0 tips the TOP edge TOWARD the viewer (top edge nearer/wider)', () => {
    // A pure latitude (X-axis) tilt. With lat > 0 the top edge tips toward the
    // viewer and the bottom recedes (file-angle convention — see the module
    // header; sample-11's lat = −30° does the opposite, top recedes). The top and
    // bottom edges stay horizontal and the quad mirrors left↔right about the
    // vertical centre line; the near (top) edge is WIDER than the far (bottom).
    const cam: CameraInput = { prst: 'perspectiveFront', rot: { lat: 25, lon: 0, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(false);
    const [tl, tr, br, bl] = q.corners.map((c) => r(c));
    // Top edge horizontal (TL.y == TR.y); bottom edge horizontal (BL.y == BR.y).
    expect(tl.y).toBeCloseTo(tr.y, 2);
    expect(bl.y).toBeCloseTo(br.y, 2);
    // Mirror about the vertical centre line: TL.x↔TR.x, BL.x↔BR.x.
    expect(tl.x).toBeCloseTo(W - tr.x, 2);
    expect(bl.x).toBeCloseTo(W - br.x, 2);
    // Top edge WIDER than bottom edge (top nears the viewer).
    const topW = tr.x - tl.x;
    const botW = br.x - bl.x;
    expect(topW).toBeGreaterThan(botW);
  });

  it('lat<0 (sample-11 slide-3 tilt) makes the TOP edge recede (narrower)', () => {
    // sample-11 slide 3 supplies lat = 330° (= −30°). The corrected convention
    // must make the top edge RECEDE here, matching the PDF ground truth
    // (top/bottom width ratio ≈ 0.82). This is the inverse of the lat>0 case and
    // is the regression guard for the axis-convention fix.
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 0, rev: 0 } };
    const q = computeScene3dQuad(cam, W, H);
    const [tl, tr, br, bl] = q.corners;
    const topW = tr.x - tl.x;
    const botW = br.x - bl.x;
    expect(topW).toBeLessThan(botW);
  });

  it('rev-only rotation is a pure in-plane rotation (affine, rigid)', () => {
    // rev spins the shape in the screen plane: the quad stays a rectangle
    // (affine), corner-to-corner distances are preserved up to the refit scale.
    // rev > 0 rotates COUNTER-CLOCKWISE on screen (file-angle convention).
    const cam: CameraInput = { prst: 'orthographicFront', rot: { lat: 0, lon: 0, rev: 30 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.isAffine).toBe(true);
    expect(q.isIdentity).toBe(false);
    // Edge lengths: opposite edges equal, adjacent edges in the original w:h
    // ratio (rigid rotation preserves the rectangle's proportions).
    const [tl, tr, br, bl] = q.corners;
    const top = Math.hypot(tr.x - tl.x, tr.y - tl.y);
    const right = Math.hypot(br.x - tr.x, br.y - tr.y);
    const bottom = Math.hypot(br.x - bl.x, br.y - bl.y);
    const left = Math.hypot(bl.x - tl.x, bl.y - tl.y);
    expect(top).toBeCloseTo(bottom, 1);
    expect(left).toBeCloseTo(right, 1);
    expect(top / right).toBeCloseTo(W / H, 2);
    // Quad centroid stays at the bbox centre.
    const cx = (tl.x + tr.x + br.x + bl.x) / 4;
    const cy = (tl.y + tr.y + br.y + bl.y) / 4;
    expect(cx).toBeCloseTo(W / 2, 3);
    expect(cy).toBeCloseTo(H / 2, 3);
    // Direction: rev > 0 is counter-clockwise on screen, so the top-right corner
    // rises above the top-left corner (TR.y < TL.y in y-down coords).
    expect(tr.y).toBeLessThan(tl.y);
  });

  it('fov override widens the perspective foreshortening', () => {
    const base: CameraInput = { prst: 'perspectiveFront', rot: { lat: 30, lon: 0, rev: 0 } };
    const wide: CameraInput = { ...base, fov: 90 };
    const qb = computeScene3dQuad(base, W, H);
    const qw = computeScene3dQuad(wide, W, H);
    // Wider FOV → stronger near/far size difference → bigger top-vs-bottom width gap.
    const gap = (q: ReturnType<typeof computeScene3dQuad>): number => {
      const topW = q.corners[1].x - q.corners[0].x;
      const botW = q.corners[2].x - q.corners[3].x;
      return Math.abs(topW - botW);
    };
    expect(gap(qw)).toBeGreaterThan(gap(qb));
  });

  it('ignores FOV for a non-perspective camera, including extrusion displacement', () => {
    // MS-OI29500 §2.1.1270: an authored FOV must not change a parallel view.
    const base: CameraInput = {
      prst: 'orthographicFront', rot: { lat: 25, lon: 35, rev: 17 },
    };
    const narrow = { ...base, fov: 5 };
    const wide = { ...base, fov: 115 };
    expect(computeScene3dQuad(narrow, W, H)).toEqual(computeScene3dQuad(wide, W, H));
    expect(computeDepthOffset(narrow, W, H, 40)).toEqual(computeDepthOffset(wide, W, H, 40));
  });

  it('camera zoom changes the projected extent instead of being cancelled', () => {
    const base: CameraInput = {
      prst: 'perspectiveFront',
      rot: { lat: 20, lon: 10, rev: 0 },
    };
    const zoomed: CameraInput = { ...base, zoom: 2 };
    const extent = (q: ReturnType<typeof computeScene3dQuad>): number => {
      const xs = q.corners.map((c) => c.x);
      return Math.max(...xs) - Math.min(...xs);
    };
    expect(extent(computeScene3dQuad(zoomed, W, H)))
      .toBeCloseTo(2 * extent(computeScene3dQuad(base, W, H)), 6);
  });

  it('does not refit a tilted projection into the original 2-D box', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 347 } };
    const q = computeScene3dQuad(cam, W, H);
    expect(q.corners.some((c) => c.x < 0 || c.x > W || c.y < 0 || c.y > H))
      .toBe(true);
  });
});

describe('computeDepthOffset (extrusion side-wall direction, §20.1.5.12)', () => {
  it('is ~zero for a face-on camera (−Z projects straight into the screen)', () => {
    const cam: CameraInput = { prst: 'perspectiveFront' };
    const o = computeDepthOffset(cam, W, H, 20);
    expect(Math.hypot(o.x, o.y)).toBeLessThan(0.5);
  });

  it('is exactly zero for orthographicFront (parallel, no tilt)', () => {
    const cam: CameraInput = { prst: 'orthographicFront' };
    const o = computeDepthOffset(cam, W, H, 20);
    expect(o.x).toBeCloseTo(0, 6);
    expect(o.y).toBeCloseTo(0, 6);
  });

  it('reveals a side wall when the camera is tilted (non-zero offset)', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 0 } };
    const o = computeDepthOffset(cam, W, H, 40);
    expect(Math.hypot(o.x, o.y)).toBeGreaterThan(1);
  });

  it('scales with the extrusion depth', () => {
    const cam: CameraInput = { prst: 'perspectiveRelaxed', rot: { lat: 330, lon: 20, rev: 0 } };
    const a = computeDepthOffset(cam, W, H, 20);
    const b = computeDepthOffset(cam, W, H, 40);
    expect(Math.hypot(b.x, b.y)).toBeGreaterThan(Math.hypot(a.x, a.y));
  });
});

describe('modern camera families', () => {
  it.each([
    ['perspectiveRelaxed', 45],
    ['perspectiveHeroicExtremeLeftFacing', 80],
    ['perspectiveHeroicExtremeRightFacing', 80],
  ] as const)('%s uses its omitted lens, while an explicit lens takes precedence', (prst, fov) => {
    const camera: CameraInput = { prst, rot: { lat: 25, lon: 35, rev: 17 } };
    expect(computeScene3dQuad(camera, W, H))
      .toEqual(computeScene3dQuad({ ...camera, fov }, W, H));
    expect(computeDepthOffset(camera, W, H, 40))
      .toEqual(computeDepthOffset({ ...camera, fov }, W, H, 40));
    expect(computeScene3dQuad({ ...camera, fov: 15 }, W, H))
      .toEqual(computeScene3dQuad({ ...camera, prst: 'perspectiveFront', fov: 15 }, W, H));
  });

  // Independent synthetic Windows PowerPoint rectangle controls, measured in
  // points. Corners are unlabelled; match the point set within the conservative
  // 0.75pt PDF extraction uncertainty. These cover distinct orientation classes.
  it.each([
    ['isometricLeftDown', [[13.63, -13.987513], [81.63, 25.041232], [81.63, 77.341319], [13.63, 38.297]]],
    ['isometricOffAxis4Bottom', [[-9.399058, 29.3229], [76.777613, 16.348932], [104.647141, 34.020438], [18.485515, 46.989077]]],
    ['perspectiveAboveLeftFacing', [[13.900154, -14.375288], [98.436788, 30.355168], [80.793806, 76.730822], [-0.423556, 32.856135]]],
    ['perspectiveRelaxed', [[0.766264, 11.92], [94.515514, 11.92], [96.659306, 52.42], [-1.397547, 52.42]]],
  ] as const)('%s supplies the omitted orientation', (prst, observed) => {
    const { corners } = computeScene3dQuad({ prst, unitsPerInch: 72 }, 95.76, 63.84);
    for (const [x, y] of observed) {
      expect(Math.min(...corners.map((p) => Math.hypot(p.x - x, p.y - y)))).toBeLessThan(0.75);
    }
  });

  it('explicit rotation replaces the isometric preset orientation', () => {
    const camera = { prst: 'isometricOffAxis4Bottom', rot: { lat: 0, lon: 0, rev: 0 } };
    expect(computeScene3dQuad(camera, W, H).isIdentity).toBe(true);
  });

  it('modern oblique uses parallel directional shear after rotation, including depth', () => {
    const camera = { prst: 'obliqueTopLeft', rot: { lat: 0, lon: 30, rev: 0 }, fov: 115 };
    const { corners, isAffine } = computeScene3dQuad(camera, 200, 100);
    // The right corners rotate to x=100*cos(30), z=50. The 0.3 shear
    // distributes z equally over the diagonal, independently of the FOV.
    expect(corners[1].x).toBeCloseTo(100 + 100 * Math.cos(Math.PI / 6) + 15 / Math.sqrt(2), 8);
    expect(corners[1].y).toBeCloseTo(15 / Math.sqrt(2), 8);
    expect(isAffine).toBe(true);
    expect(computeScene3dQuad({ ...camera, fov: 5 }, 200, 100).corners).toEqual(corners);
    expect(computeDepthOffset({ prst: 'obliqueLeft' }, 200, 100, 20)).toEqual({ x: -6, y: 0 });
    expect(computeDepthOffset({ prst: 'obliqueRight' }, 200, 100, 20).x).toBeCloseTo(6, 8);
  });

  it('routes depth geometry even when the default oblique front plane is identity', () => {
    for (const prst of ['obliqueTop', 'obliqueTopLeft']) {
      const camera = { prst };
      expect(isScene3dNonIdentity(camera, 200, 100, 0)).toBe(false);
      expect(isScene3dNonIdentity(camera, 200, 100, 60)).toBe(true);
      expect(isScene3dNonIdentity(camera, undefined, undefined, 60)).toBe(true);
    }
    expect(isScene3dNonIdentity({ prst: 'orthographicFront' }, 200, 100, 60)).toBe(false);
  });

  it('does not route a hidden inset back plane through another front-face rasterization', () => {
    const camera = { prst: 'perspectiveFront', unitsPerInch: 72, fov: 90 };
    expect(computeScene3dQuad(camera, 288, 360, 2).isIdentity).toBe(false);
    // A convex extrusion receding straight behind the unchanged opaque front
    // cannot expose a wall or back face, even when its inset is measurable.
    expect(isScene3dNonIdentity(camera, 288, 360, 2)).toBe(false);
  });

  it('modern perspective distance is physical and independent of shape size/aspect', () => {
    const camera = { prst: 'perspectiveFront', unitsPerInch: 72, fov: 90, rot: { lat: 0, lon: 30, rev: 0 } };
    const wide = computeScene3dQuad(camera, 95.76, 63.84);
    const tall = computeScene3dQuad(camera, 95.76, 191.52);
    // Independent ray-plane result: x=47.88+41.4652963332/(1-23.94/452.16).
    expect(wide.corners[1].x).toBeCloseTo(91.6634486713, 7);
    expect(tall.corners[1].x).toBeCloseTo(wide.corners[1].x, 8);
    // Doubling physical width changes depth relative to the fixed camera.
    // Right edge: x=95.76+82.9305926664/(1-47.88/452.16).
    expect(computeScene3dQuad(camera, 191.52, 127.68).corners[1].x)
      .toBeCloseTo(188.5122924212, 7);
    // A different width must not change the centre's depth projection.
    expect(computeDepthOffset(camera, 191.52, 127.68, 20))
      .toEqual(computeDepthOffset(camera, 95.76, 63.84, 20));
  });

  it('keeps CSS and device projections consistent under uniform scaling', () => {
    const camera = { prst: 'perspectiveHeroicExtremeLeftFacing', zoom: 1.5, rot: { lat: 25, lon: 35, rev: 17 } };
    const a = computeScene3dQuad(camera, W, H);
    const deviceCamera = { ...camera, unitsPerInch: 96 * 4 };
    const b = computeScene3dQuad(deviceCamera, W * 4, H * 4);
    a.corners.forEach((p, i) => {
      expect(b.corners[i].x).toBeCloseTo(p.x * 4, 8);
      expect(b.corners[i].y).toBeCloseTo(p.y * 4, 8);
    });
    const da = computeDepthOffset(camera, W, H, 20);
    const db = computeDepthOffset(deviceCamera, W * 4, H * 4, 80);
    expect(db.x).toBeCloseTo(da.x * 4, 8);
    expect(db.y).toBeCloseTo(da.y * 4, 8);
  });

  it.each(['legacyObliqueTopLeft', 'legacyPerspectiveBottomRight'])('retains %s fallback behavior', (prst) => {
    const camera = { prst, fov: 115, zoom: 1.5, rot: { lat: 25, lon: 35, rev: 17 } };
    expect(computeScene3dQuad(camera, W, H))
      .toEqual(computeScene3dQuad({ ...camera, prst: 'orthographicFront' }, W, H));
    expect(computeDepthOffset(camera, W, H, 40))
      .toEqual(computeDepthOffset({ ...camera, prst: 'orthographicFront' }, W, H, 40));
  });
});

describe('perspective scene boundary safety', () => {
  it.each([.99, 1, 1.01])('keeps a flat scene convex at %s times the camera distance', ratio => {
    const angle = Math.PI / 3;
    const base = 6.28 * 96;
    const w = 2 * base * ratio / Math.sin(angle), h = 96;
    const near = w / 2 * Math.sin(angle);
    // Library fallback retains a two-percent gap to the rotated z-volume,
    // moving only along view Z and retaining the centre-plane framing.
    const distance = Math.max(base, near + .02 * 2 * near);
    const q = computeScene3dQuad({ prst: 'perspectiveFront', fov: 90,
      rot: { lat: 0, lon: 60, rev: 0 } }, w, h);
    expect(q.corners[1].x).toBeCloseTo(w / 2 + w / 2 * Math.cos(angle) / (1 - near / distance), 7);
    const turns = q.corners.map((a, i, p) => {
      const b = p[(i + 1) % 4], c = p[(i + 2) % 4];
      return (b.x - a.x) * (c.y - b.y) - (b.y - a.y) * (c.x - b.x);
    });
    expect(turns.every(v => v > 0)).toBe(true);
  });

  it('uses the complete solid for the face, rear surface and centre displacement at every raster scale', () => {
    // lon=120 turns an 8-inch extrusion toward the camera. The face alone is
    // in front of it; only the rear corners force the distance adjustment.
    for (const units of [72, 96, 384]) {
      const w = 8 * units, h = units, depth = 8 * units;
      const camera = { prst: 'perspectiveFront', fov: 90, unitsPerInch: units,
        sceneDepth: depth, rot: { lat: 0, lon: 120, rev: 0 } };
      const extent = 2 * Math.sqrt(3) * units;
      const distance = extent + 4 * units + .02 * (2 * extent + 4 * units);
      const face = computeScene3dQuad(camera, w, h);
      const back = computeScene3dQuad(camera, w, h, depth);
      expect(face.corners[1].x).toBeCloseTo(w / 2 - 2 * units / (1 - extent / distance), 7);
      expect(back.corners[1].x).toBeCloseTo(w / 2 + (4 * Math.sqrt(3) - 2) * units / (1 - (extent + 4 * units) / distance), 7);
      const offset = computeDepthOffset(camera, w, h, depth);
      expect(offset.x).toBeCloseTo(4 * Math.sqrt(3) * units / (1 - 4 * units / distance), 7);
      expect(offset.y).toBeCloseTo(0, 7);
    }
  });

  it('keeps endpoint lenses and combined rotation finite without reversing any projected edge', () => {
    for (const fov of [0, 180]) {
      const q = computeScene3dQuad({ prst: 'perspectiveFront', fov,
        rot: { lat: 50, lon: 65, rev: 17 } }, 2000, 1000);
      expect(q.corners.every(p => Number.isFinite(p.x) && Number.isFinite(p.y))).toBe(true);
      const turns = q.corners.map((a, i, p) => {
        const b = p[(i + 1) % 4], c = p[(i + 2) % 4];
        return (b.x - a.x) * (c.y - b.y) - (b.y - a.y) * (c.x - b.x);
      });
      expect(turns.every(v => v > 0)).toBe(true);
    }
  });
});
