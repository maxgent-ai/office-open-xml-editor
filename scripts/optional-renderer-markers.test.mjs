import assert from 'node:assert/strict';
import test from 'node:test';
import { rolldown } from 'rolldown';
import { optionalRendererMarkers } from './optional-renderer-markers.mjs';

// Exercise the real retained-module table, including an imported but dead
// implementation, rather than manufacturing renderChunk metadata.
test('keeps boundary markers only for retained optional code without debug comments', async () => {
  const chart = '/fixture/packages/core/src/chart/chart-ex-renderer.ts';
  const tiff = '/fixture/packages/core/src/image/tiff.ts';
  const bundle = await rolldown({
    input: '/fixture/entry.js',
    experimental: { attachDebugInfo: 'none' },
    plugins: [{
      name: 'fixture',
      resolveId: id => id,
      load(id) {
        if (id === chart) return 'export const chartEx = /* @__PURE__ */ Object.freeze({ render() { return 42; } });';
        if (id === tiff) return 'export const tiff = { render() { return 99; } };';
        return `import { chartEx } from '${chart}'; import { tiff } from '${tiff}'; export { chartEx };`;
      },
    }, optionalRendererMarkers()],
  });
  try {
    const { output } = await bundle.generate({ format: 'es' });
    const code = output[0].code;
    assert.ok(code.includes('optional-renderer: packages/core/src/chart/chart-ex-renderer.ts'));
    assert.ok(!code.includes('packages/core/src/image/tiff.ts'));
    assert.ok(!code.includes('#region'));
    assert.ok(code.includes('@__PURE__'), 'keep consumer tree-shaking annotations');
    const result = await import(`data:text/javascript,${encodeURIComponent(code)}`);
    assert.equal(result.chartEx.render(), 42);
  } finally {
    await bundle.close();
  }
});
