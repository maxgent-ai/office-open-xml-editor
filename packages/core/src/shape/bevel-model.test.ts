import { prepareLighting, rigVector, shadeSurface } from './bevel-lighting';
import { describe, expect, it } from 'vitest';
import {
  applyBevelShading,
  applyExtrusion,
  bevelHeightProfile,
  lightDirFromRig,
  materialClass,
  type BevelCtx,
} from './bevel-shading';

function render(
  material = 'matte',
  rig = 'threePt',
  dir = 't',
  rot?: { lat: number; lon: number; rev: number },
) {
  const data = new Uint8ClampedArray(40 * 40 * 4);
  for (let i = 0; i < data.length; i += 4) data.set([80, 100, 140, 255], i);
  const ctx: BevelCtx = {
    canvas: { width: 40, height: 40 },
    getImageData: () =>
      ({ data: data.slice(), width: 40, height: 40 }) as ImageData,
    putImageData: (image) => {
      data.set(image.data);
    },
  };
  applyBevelShading(ctx, {
    widthPx: 8,
    heightPx: 8,
    prst: 'circle',
    material: materialClass(material),
    light: lightDirFromRig(rig, dir, rot),
  });
  return data;
}

describe('documented DrawingML bevel model', () => {
  it('preserves material identity through the existing caller adapter', () => {
    for (const name of [
      'legacyMatte',
      'legacyPlastic',
      'legacyMetal',
      'legacyWireframe',
      'matte',
      'plastic',
      'metal',
      'warmMatte',
      'translucentPowder',
      'powder',
      'dkEdge',
      'softEdge',
      'clear',
      'flat',
      'softmetal',
    ]) {
      expect(materialClass(name)).toBe(name);
    }
  });
  it('reads documented points as (inward distance, height), as Office renders them', () => {
    const circle = bevelHeightProfile('circle', 10);
    // Cubic midpoint (.29139875, .70860125): a convex quarter round that is
    // already at 71% height 29% of the way in, not a cove that is still at 29%.
    expect(circle(2.9139875)).toBeCloseTo(0.70860125, 3);
    expect(bevelHeightProfile('angle', 10)(5)).toBeCloseTo(0.5, 5);
    // cross: flat step at height .1 from inward .1 to .4; its face starts at
    // inward .6 (the inward axis is not stretched).
    expect(bevelHeightProfile('cross', 10)(2.5)).toBeCloseTo(0.1, 3);
    expect(bevelHeightProfile('cross', 10)(7)).toBe(1);
    // softRound ends at height 1/3: the remainder of the band is that face.
    expect(bevelHeightProfile('softRound', 10)(10)).toBeCloseTo(1 / 3, 4);
  });
  it('keeps all documented profile envelopes finite and selects visible overhangs', () => {
    for (const name of [
      'angle',
      'artDeco',
      'circle',
      'convex',
      'coolSlant',
      'cross',
      'divot',
      'hardEdge',
      'relaxedInset',
      'riblet',
      'slope',
      'softRound',
    ]) {
      const profile = bevelHeightProfile(name, 1);
      for (let i = 0; i <= 100; i++) {
        expect(Number.isFinite(profile(i / 100))).toBe(true);
        expect(profile(i / 100)).toBeGreaterThanOrEqual(0);
        expect(profile(i / 100)).toBeLessThanOrEqual(1);
      }
    }
    // Divot folds back over inward .54–.56; the raised branch hides the low
    // one. relaxedInset peaks at inward .508 and settles to its .64 face.
    expect(bevelHeightProfile('divot', 1)(0.55)).toBeGreaterThan(0.9);
    expect(bevelHeightProfile('relaxedInset', 1)(0.508)).toBeCloseTo(1, 3);
    expect(bevelHeightProfile('relaxedInset', 1)(1)).toBeCloseTo(0.64, 5);
  });
  it('explicit rotation replaces direction, including latitude and longitude', () => {
    const rot = { lat: 25, lon: 35, rev: 15 };
    expect(render('matte', 'morning', 't', rot)).toEqual(
      render('matte', 'morning', 'r', rot),
    );
    expect(render('matte', 'morning', 't', rot)).not.toEqual(
      render('matte', 'morning', 't', { lat: 0, lon: 0, rev: 15 }),
    );
  });
  it('composes the rig rotation about its tilted central axis', () => {
    const base = lightDirFromRig('threePt', 't');
    const turned = lightDirFromRig('threePt', 't', {
      lat: 90, lon: 90, rev: 90,
    });
    // Ry(90) Rx(90) Rz(90): (x,y,z) -> (x,-z,y).
    expect(turned.x).toBeCloseTo(base.x, 10);
    expect(turned.y).toBeCloseTo(-base.z, 10);
    expect(turned.z).toBeCloseTo(base.y, 10);
  });
  it('rotates every light when shading the face, including side and back lighting', () => {
    const center = (20 * 40 + 20) * 4;
    const face = (lat: number, lon: number, rev = 0) =>
      render('matte', 'threePt', 't', { lat, lon, rev })[center + 2];
    // Independent Lambert sums of the documented threePt lights: latitude
    // +90 exposes the fill/rim, longitude +90 exposes the key/rim; -90 exposes
    // only the weak fill. No Office pixel constants or fitted gains.
    expect(face(90, 0)).toBe(138);
    expect(face(0, 90)).toBe(82);
    expect(face(0, 270)).toBe(17);
    expect(face(180, 0)).toBe(109);
    expect(face(0, 0, 90)).toBe(face(0, 0));
  });
  it('honors explicit zero over dir while retaining the omitted-rotation direction', () => {
    const zero = { lat: 0, lon: 0, rev: 0 };
    const explicit = render('matte', 'threePt', 't', zero);
    expect(render('matte', 'threePt', 'r', zero)).toEqual(explicit);
    expect(render('matte', 'threePt', 'br', zero)).toEqual(explicit);
    expect(render('matte', 'threePt', 'r')).not.toEqual(explicit);
    expect(render('matte', 'threePt', 't')).toEqual(explicit);
  });
  it('shades the flat face and retains colored rig and material responses', () => {
    const center = (20 * 40 + 20) * 4;
    expect([...render().slice(center, center + 3)]).not.toEqual([80, 100, 140]);
    expect(render('matte', 'sunset')).not.toEqual(render('matte', 'threePt'));
    expect(render('powder')).not.toEqual(render('legacyMetal'));
    expect(render('translucentPowder')[center + 3]).toBeLessThan(255);
  });
});

// Independent arithmetic expectations exercise production lighting rather than
// copying the implementation's table or checking its private representation.

describe('reference lighting arithmetic and surface visibility', () => {
  it('uses ambient-only diffuse in contrasting, including the powder color mix', () => {
    const out = new Float64Array(4);
    shadeSurface(
      out,
      80,
      100,
      140,
      255,
      0,
      0,
      1,
      prepareLighting('matte', rigVector('contrasting', 't')),
    );
    expect([...out]).toEqual([80, 100, 140, 255]);
    shadeSurface(
      out,
      80,
      100,
      140,
      255,
      0,
      0,
      1,
      prepareLighting('powder', rigVector('contrasting', 't')),
    );
    expect(out[0]).toBeCloseTo((80 + 255) / 2, 8);
    expect(out[1]).toBeCloseTo((100 + 255) / 2, 8);
  });
  it('normalizes legacy direction vectors and preserves their light energy', () => {
    const out = new Float64Array(4);
    shadeSurface(
      out,
      80,
      100,
      140,
      255,
      0,
      0,
      1,
      prepareLighting('legacyMatte', rigVector('legacyNormal1', 't')),
    );
    expect(out[0]).toBeCloseTo(80 * (0.153 + 0.671 + 0.366), 8);
    expect(out[1]).toBeCloseTo(100 * (0.153 + 0.671 + 0.366), 8);
  });
  it('uses the documented signed negative light contribution before clamping', () => {
    const out = new Float64Array(4);
    shadeSurface(
      out,
      80,
      100,
      140,
      255,
      0,
      0,
      1,
      prepareLighting('matte', rigVector('brightRoom', 't')),
    );
    // Only ambient 1.5 and the negative frontal diffuse light -.5 contribute.
    for (const [i, expected] of [80, 100, 140, 255].entries())
      expect(out[i]).toBeCloseTo(expected, 8);
  });
  it('leaves the top surface alone when asked to shade an invisible bottom face', () => {
    const data = new Uint8ClampedArray([80, 100, 140, 255]);
    const ctx: BevelCtx = {
      canvas: { width: 1, height: 1 },
      getImageData: () => ({ data }) as ImageData,
      putImageData: () => {
        throw new Error('Back surface is not in the front bitmap');
      },
    };
    applyBevelShading(ctx, {
      widthPx: 8,
      heightPx: 8,
      prst: 'circle',
      material: 'metal',
      light: lightDirFromRig('threePt', 't'),
      bottom: true,
    });
    expect([...data]).toEqual([80, 100, 140, 255]);
  });
  it('lights an exposed bottom profile with its outward normal in the camera frame', () => {
    const w = 40, data = new Uint8ClampedArray(w * w * 4);
    for (let i = 0; i < w * w; i++) data.set([80, 100, 140, 255], i * 4);
    const ctx: BevelCtx = {
      canvas: { width: w, height: w },
      getImageData: () => ({ data: data.slice(), width: w, height: w }) as ImageData,
      putImageData: image => { data.set(image.data); },
    };
    const input = { widthPx: 8, heightPx: 8, prst: 'angle', material: 'matte' as const,
      light: { x: 0, y: 0, z: 1 },
      camera: { prst: 'orthographicFront', rot: { lat: 0, lon: 180, rev: 0 } } };
    applyBevelShading(ctx, { ...input, bottom: true });
    const center = (20 * w + 20) * 4, rim = (20 * w + 2) * 4;
    expect([...data.slice(center, center + 4)]).toEqual([80, 100, 140, 255]);
    // Angle bevel has a 45-degree normal; frontal Lambert response is 1/sqrt(2).
    expect(data[rim]).toBe(57);
    const back = data.slice();
    applyBevelShading(ctx, input); // hidden front cannot relight the exposed back
    expect(data).toEqual(back);
  });

  it('draws wire rims and back edges without a filled face or a material response', () => {
    const w = 48, data = new Uint8ClampedArray(w * w * 4);
    for (let y = 8; y < 32; y++) for (let x = 8; x < 32; x++)
      data.set([80, 100, 140, 255], (y * w + x) * 4);
    const ctx: BevelCtx = {
      canvas: { width: w, height: w },
      getImageData: () => ({ data: data.slice(), width: w, height: w }) as ImageData,
      putImageData: image => { data.set(image.data); },
    };
    applyBevelShading(ctx, { widthPx: 4, heightPx: 4, prst: 'angle',
      material: 'legacyWireframe', light: { x: 0, y: 0, z: -1 },
      wireframeDepth: { x: 8, y: 4 } });
    const pixel = (x: number, y: number) => [...data.slice((y * w + x) * 4, (y * w + x) * 4 + 4)];
    expect(pixel(24, 24)[3]).toBe(0);
    expect(pixel(8, 20)).toEqual([80, 100, 140, 255]);
    expect(pixel(12, 20)[3]).toBeGreaterThan(0); // inner profile rim
    expect(pixel(39, 24)).toEqual([80, 100, 140, 255]); // translated back rim
  });

  it.each([125, 128])('keeps bevel and extrusion geometry at source alpha %i', alpha => {
    const w = 32, data = new Uint8ClampedArray(w * w * 4);
    for (let y = 4; y < 24; y++) for (let x = 4; x < 24; x++)
      data.set([80, 100, 140, alpha], (y * w + x) * 4);
    const ctx: BevelCtx = {
      canvas: { width: w, height: w },
      getImageData: () => ({ data: data.slice(), width: w, height: w }) as ImageData,
      putImageData: image => { data.set(image.data); },
    };
    applyBevelShading(ctx, { widthPx: 4, heightPx: 4, prst: 'angle',
      material: 'matte', light: { x: 0, y: 0, z: 1 } });
    const pixel = (x: number, y: number) => [...data.slice((y * w + x) * 4, (y * w + x) * 4 + 4)];
    // Angle's 45-degree normal gives 80/sqrt(2) at the lip, at either
    // opacity. Lighting preserves source alpha rather than flattening the lip.
    expect(pixel(5, 14)).toEqual([57, 71, 99, alpha]);
    const face = pixel(14, 14);
    applyExtrusion(ctx, { offsetX: 4, offsetY: 0, rgb: [20, 30, 40] });
    expect(pixel(14, 14)).toEqual(face); // translucent face still occludes wall
    expect(pixel(26, 14)).toEqual([20, 30, 40, 255]);
  });

  it('bounds wireframe work to a noisy local raster despite an enormous depth', () => {
    const w = 96, data = new Uint8ClampedArray(w * w * 4);
    for (let y = 0; y < w; y++) for (let x = 0; x < w; x++)
      if ((Math.floor(x / 2) + Math.floor(y / 2)) % 2 === 0)
        data.set([80, 100, 140, 255], (y * w + x) * 4);
    const ctx: BevelCtx = {
      canvas: { width: w, height: w },
      getImageData: () => ({ data: data.slice(), width: w, height: w }) as ImageData,
      putImageData: image => { data.set(image.data); },
    };
    applyBevelShading(ctx, { widthPx: 4, heightPx: 4, prst: 'angle',
      material: 'legacyWireframe', light: { x: 0, y: 0, z: 1 },
      wireframeDepth: { x: 1e12, y: 1e12 } });
    // Depth edges leave the raster rather than requesting a depth-sized buffer
    // or iterating to their off-canvas endpoints. The visible wire survives.
    expect([...data.slice(-4)]).toEqual([80, 100, 140, 255]);
  });

  it('defaults unknown preset names safely, including object property names', () => {
    expect(lightDirFromRig('constructor', '__proto__')).toEqual(
      lightDirFromRig('threePt', 't'),
    );
    expect(materialClass('__proto__')).toBe('matte');
    expect(bevelHeightProfile('constructor', 10)(5)).toEqual(
      bevelHeightProfile('circle', 10)(5),
    );
  });
});
