// The built-bundle boundary checks need four ownership markers, not Rolldown's
// per-module debug comments. Derive them from retained modules so tree-shaken
// imports do not look like a leaked implementation. No source paths outside
// this fixed public allowlist are emitted.
const implementations = [
  'packages/core/src/chart/chart-ex-renderer.ts',
  'packages/core/src/chart/three-d-renderer.ts',
  'packages/core/src/chart/region-map-natural-earth.generated.ts',
  'packages/core/src/image/tiff.ts',
];

export function optionalRendererMarkers() {
  return {
    name: 'optional-renderer-markers',
    renderChunk(code, chunk) {
      const retained = Object.entries(chunk.modules)
        .filter(([, module]) => module.renderedLength > 0)
        .map(([id]) => id.replaceAll('\\', '/').split('?')[0]);
      const markers = implementations.filter(path => retained.some(id => id.endsWith(`/${path}`)));
      if (markers.length === 0) return null;
      return { code: `${markers.map(path => `// optional-renderer: ${path}`).join('\n')}\n${code}`, map: null };
    },
  };
}
