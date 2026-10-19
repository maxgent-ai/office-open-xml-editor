import { describe, expect, it } from 'vitest';
import { createGeometryPath, type GeometryPathSink } from './path-data';
import { buildShapePath } from './preset';
import { buildPresetGeometryPath } from './preset-geometry';

function trace(build: (sink: GeometryPathSink) => unknown) {
  const path = createGeometryPath();
  build(path.sink);
  return path.commands;
}

describe('buildShapePath fallback', () => {
  it('emits the implicit box for rect and for names the spec engine does not define', () => {
    for (const name of ['rect', 'notAShape']) {
      expect(trace((sink) => buildShapePath(sink, name, 1, 2, 30, 40))).toEqual([
        { op: 'rect', args: [1, 2, 30, 40] },
      ]);
    }
  });

  it('routes spec-defined names and legacy aliases through the ECMA-376 engine', () => {
    for (const [name, spec] of [['homePlate', 'homeplate'], ['star', 'star5'], ['oval', 'ellipse']]) {
      expect(trace((sink) => buildShapePath(sink, name, 0, 0, 120, 60, 30000)))
        .toEqual(trace((sink) => buildPresetGeometryPath(sink, spec, 0, 0, 120, 60, [30000, null, null, null])));
    }
  });
});
