import { describe, expect, it } from 'vitest';
import {
  applyLinearTintOrShade,
  isObservedAutomaticSurfaceCamera,
  legacyPattern2Color,
  surfaceMaterialFactor,
  surfacePerspectiveTangentGain,
} from './material-color.js';
import { automaticSurfaceMajorUnit } from './axis-scale.js';

const ACCENTS = ['156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E'];

describe('legacy Pattern 2 generated colours', () => {
  it('matches the shared parser linear-sRGB transform parity fixtures', () => {
    expect(applyLinearTintOrShade('#156082', -0.35)).toBe('#0F4E6A');
    expect(applyLinearTintOrShade('#156082', 0)).toBe('#156082');
    expect(applyLinearTintOrShade('#156082', 0.35)).toBe('#A1AFBB');
  });
  it('keeps a partial first set at the base accents', () => {
    expect(Array.from({ length: 4 }, (_, index) =>
      legacyPattern2Color(ACCENTS, index, 4, 2)))
      .toEqual(['#156082', '#E97132', '#196B24', '#0F9ED5']);
  });

  it('keeps registered set transforms independent of the final object count', () => {
    expect(Array.from({ length: 7 }, (_, index) =>
      legacyPattern2Color(ACCENTS, index, 7, 2)))
      .toEqual([
        '#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E',
        applyLinearTintOrShade('#156082', -0.4),
      ]);
    expect(legacyPattern2Color(ACCENTS, 6, 13, 2))
      .toBe(legacyPattern2Color(ACCENTS, 6, 7, 2));
    expect(legacyPattern2Color(ACCENTS, 48, 49, 2)).toBeNull();
  });
});

describe('surface automatic material', () => {
  it('is winding-invariant and bounded to the surface compatibility range', () => {
    const lit = surfaceMaterialFactor({ x: 0.2, y: 0.4, z: 0.9 });
    expect(surfaceMaterialFactor({ x: -0.2, y: -0.4, z: -0.9 })).toBeCloseTo(lit, 12);
    expect(lit).toBeGreaterThan(1);
    expect(surfaceMaterialFactor({ x: 0.2, y: -0.4, z: 0.1 })).toBeGreaterThanOrEqual(0.48);
  });

  it('lights camera-space upper-right facets more strongly than their mirrors', () => {
    expect(surfaceMaterialFactor({ x: 0.4, y: 0, z: 0.92 }))
      .toBeGreaterThan(surfaceMaterialFactor({ x: -0.4, y: 0, z: 0.92 }));
    expect(surfaceMaterialFactor({ x: 0, y: 0.4, z: 0.92 }))
      .toBeGreaterThan(surfaceMaterialFactor({ x: 0, y: -0.4, z: 0.92 }));
  });
});

describe('surface automatic camera compatibility', () => {
  const observedCamera = {
    rotationX: 15,
    rotationY: 20,
    rightAngleAxes: false,
    perspective: 30,
  };

  it('uses the observed perspective gain only for the effective omitted-view camera', () => {
    expect(isObservedAutomaticSurfaceCamera(observedCamera)).toBe(true);
    expect(surfacePerspectiveTangentGain(observedCamera)).toBe(2);
  });

  it.each([
    { ...observedCamera, rotationX: 30 },
    { ...observedCamera, rotationY: 45 },
    { ...observedCamera, rightAngleAxes: true },
    { ...observedCamera, perspective: 20 },
  ])('keeps authored camera projections outside the observed boundary', camera => {
    expect(isObservedAutomaticSurfaceCamera(camera)).toBe(false);
    expect(surfacePerspectiveTangentGain(camera)).toBe(1);
  });
});

describe('surface automatic major unit', () => {
  it('uses projected axis length and the compact edge-on floor', () => {
    expect(automaticSurfaceMajorUnit(0, 30, 220)).toBe(5);
    expect(automaticSurfaceMajorUnit(0, 40, 136)).toBe(10);
    expect(automaticSurfaceMajorUnit(10, 90, 0)).toBe(20);
    expect(automaticSurfaceMajorUnit(1, 32, 220)).toBe(5);
  });
});
