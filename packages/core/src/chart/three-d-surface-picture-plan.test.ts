import { describe, expect, it } from 'vitest';
import {
  planChartThreeDSurfacePicture,
  surfacePictureFaceIsEnabled,
  surfacePictureFaceRepetitions,
  surfacePictureFaceUsesValueAxis,
} from './three-d-surface-picture-plan.js';

const fill = {
  fillType: 'image' as const,
  imagePath: 'xl/media/surface.png',
  mimeType: 'image/png',
  stretch: true,
};

describe('positive-thickness CT_Surface picture faces', () => {
  it('accepts bounded source crops and observed outsets only for stretch', () => {
    expect(planChartThreeDSurfacePicture({
      ...fill,
      srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
    }, {
      thicknessPercent: 0,
      pictureOptions: { applyToFront: true, pictureFormat: 'stretch' },
    }, 'backWall', 10)).toEqual({ mode: 'stretch', repetitions: 1, slabFaces: undefined });
    expect(planChartThreeDSurfacePicture({
      ...fill,
      srcRect: { l: -0.25, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toEqual({
      mode: 'stretch', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture({
      ...fill,
      srcRect: { l: 0.6, t: 0, r: 0.4, b: 0 },
    }, undefined, 'backWall', 10)).toBeNull();
    expect(planChartThreeDSurfacePicture({
      ...fill,
      srcRect: { l: -0.25, t: 0, r: 0, b: 0 },
    }, {
      pictureOptions: {
        pictureFormat: 'stackScale',
        pictureStackUnit: 2,
      },
    }, 'backWall', 10)).toBeNull();
  });

  it('accepts bounded stretch destination insets and outsets only', () => {
    expect(planChartThreeDSurfacePicture({
      ...fill,
      fillRect: { l: 0.25, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toEqual({
      mode: 'stretch', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture({
      ...fill,
      fillRect: { l: -0.25, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toEqual({
      mode: 'stretch', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture({
      ...fill,
      fillRect: { l: 0.6, t: 0, r: 0.4, b: 0 },
    }, undefined, 'backWall', 10)).toBeNull();
    expect(planChartThreeDSurfacePicture({
      ...fill,
      fillRect: { l: 0.25, t: 0, r: 0, b: 0 },
    }, {
      pictureOptions: {
        pictureFormat: 'stackScale',
        pictureStackUnit: 2,
      },
    }, 'backWall', 10)).toBeNull();
  });

  it.each([
    ['front', { applyToFront: true, applyToSides: false, applyToEnd: false }, [0]],
    ['sides', { applyToFront: false, applyToSides: true, applyToEnd: false }, [3, 5]],
    ['end', { applyToFront: false, applyToSides: false, applyToEnd: true }, [2, 4]],
  ] as const)('maps %s to its independently authored slab face class', (_name, flags, faces) => {
    const plan = planChartThreeDSurfacePicture(fill, {
      thicknessPercent: 25,
      pictureOptions: { ...flags, pictureFormat: 'stretch' },
    }, 'backWall', 10);
    expect(plan).not.toBeNull();
    if (!plan) throw new Error('picture plan not built');
    expect(Array.from({ length: 6 }, (_, index) => surfacePictureFaceIsEnabled(plan, index)))
      .toEqual(Array.from({ length: 6 }, (_, index) => faces.some(face => face === index)));
  });

  it('treats omitted face flags as enabled and repeats the observed thick stackScale faces', () => {
    const stretch = planChartThreeDSurfacePicture(fill, {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stretch' },
    }, 'sideWall', 10);
    expect(stretch).not.toBeNull();
    if (!stretch) throw new Error('picture plan not built');
    expect(Array.from({ length: 6 }, (_, index) => surfacePictureFaceIsEnabled(stretch, index)))
      .toEqual([true, false, true, true, true, true]);
    const stackScale = planChartThreeDSurfacePicture(fill, {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stackScale', pictureStackUnit: 2 },
    }, 'sideWall', 10);
    expect(stackScale).toEqual({
      mode: 'stackScale',
      repetitions: 5,
      stackUnit: 2,
      slabFaces: { front: true, sides: true, end: true },
    });
    if (!stackScale) throw new Error('stackScale picture plan not built');
    expect(Array.from({ length: 6 }, (_, index) =>
      surfacePictureFaceRepetitions(stackScale, index))).toEqual([5, 0, 1, 5, 1, 5]);
    expect(Array.from({ length: 6 }, (_, index) =>
      surfacePictureFaceUsesValueAxis(stackScale, index)))
      .toEqual([true, false, false, true, false, true]);
    const oneRepetition = planChartThreeDSurfacePicture(fill, {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stackScale', pictureStackUnit: 20 },
    }, 'sideWall', 10);
    if (!oneRepetition) throw new Error('one-repetition picture plan not built');
    expect(surfacePictureFaceRepetitions(oneRepetition, 0)).toBe(1);
    expect(surfacePictureFaceUsesValueAxis(oneRepetition, 0)).toBe(true);
    expect(surfacePictureFaceRepetitions(oneRepetition, 2)).toBe(1);
    expect(surfacePictureFaceUsesValueAxis(oneRepetition, 2)).toBe(false);
    expect(planChartThreeDSurfacePicture(fill, {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stackScale', pictureStackUnit: 2 },
    }, 'floor', 10)).toEqual({
      mode: 'stretch',
      repetitions: 1,
      slabFaces: { front: true, sides: true, end: true },
    });
  });

  it('bounds thick stackScale work across every enabled slab face', () => {
    const frontOnly = {
      thicknessPercent: 25,
      pictureOptions: {
        applyToFront: true,
        applyToSides: false,
        applyToEnd: false,
        pictureFormat: 'stackScale' as const,
        pictureStackUnit: 2,
      },
    };
    expect(planChartThreeDSurfacePicture(fill, frontOnly, 'backWall', 8_192)).not.toBeNull();
    expect(() => planChartThreeDSurfacePicture(fill, frontOnly, 'backWall', 8_194)).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));

    const allFaces = {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stackScale' as const, pictureStackUnit: 2 },
    };
    expect(planChartThreeDSurfacePicture(fill, allFaces, 'backWall', 2_728)).not.toBeNull();
    expect(() => planChartThreeDSurfacePicture(fill, allFaces, 'backWall', 2_730)).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
  });

  it('retains plain stack only for the observed identity rectangle boundary', () => {
    const surface = {
      thicknessPercent: 0,
      pictureOptions: { pictureFormat: 'stack' as const },
    };
    expect(planChartThreeDSurfacePicture(fill, surface, 'backWall', 10)).toEqual({
      mode: 'stack', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture(fill, {
      ...surface,
      thicknessPercent: 25,
    }, 'backWall', 10)).toEqual({
      mode: 'stack',
      repetitions: 1,
      slabFaces: { front: true, sides: true, end: true },
    });
    expect(planChartThreeDSurfacePicture({
      ...fill,
      srcRect: { l: 0.1, t: 0, r: 0, b: 0 },
    }, surface, 'backWall', 10)).toBeNull();
    expect(planChartThreeDSurfacePicture({
      ...fill,
      fillRect: { l: -0.1, t: 0, r: 0, b: 0 },
    }, surface, 'backWall', 10)).toBeNull();
  });

  it('retains an explicit DrawingML tile grid on planar and thick surfaces', () => {
    const tiled = {
      ...fill,
      stretch: false,
      dpi: 96,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
    };
    expect(planChartThreeDSurfacePicture(tiled, {
      thicknessPercent: 0,
      pictureOptions: { applyToFront: true, pictureFormat: 'stretch' },
    }, 'backWall', 10)).toEqual({ mode: 'tile', repetitions: 1, slabFaces: undefined });
    expect(planChartThreeDSurfacePicture(tiled, {
      thicknessPercent: 25,
      pictureOptions: { pictureFormat: 'stretch' },
    }, 'backWall', 10)).toEqual({
      mode: 'tile',
      repetitions: 1,
      slabFaces: { front: true, sides: true, end: true },
    });
    expect(planChartThreeDSurfacePicture({
      ...tiled,
      srcRect: { l: 0.1, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toEqual({
      mode: 'tile', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture({
      ...tiled,
      srcRect: { l: -0.1, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toEqual({
      mode: 'tile', repetitions: 1, slabFaces: undefined,
    });
    expect(planChartThreeDSurfacePicture({
      ...tiled,
      fillRect: { l: 0.1, t: 0, r: 0, b: 0 },
    }, undefined, 'backWall', 10)).toBeNull();
  });
});
