import { expect, it } from 'vitest';
import { ScrollZoomController } from './scroll-zoom-controller';

function fixture() {
  let count = 0;
  let base = 1;
  let fitWidth = 200;
  const events: string[] = [];
  const host = { scrollTop: 60, scrollLeft: 20, clientHeight: 100, clientWidth: 200 } as HTMLDivElement;
  const spacer = { style: { height: '', width: '' }, offsetWidth: 1000 } as HTMLDivElement;
  const zoom = new ScrollZoomController({
    scrollHost: () => host, spacer: () => spacer, count: () => count,
    zoomMin: () => 0.1, zoomMax: () => 8, baseScale: () => base,
    fitWidthPx: () => fitWidth,
    fitContentSize: () => ({ width: 100, height: 200 }),
    indexAt: (y) => y < 20 + 100 * zoom.scale ? 0 : 1,
    offset: (index) => index === 0 ? 10 : 20 + 100 * zoom.scale,
    height: (index) => (index === 0 ? 100 : 200) * zoom.scale,
    totalHeight: () => 30 + 300 * zoom.scale,
    recomputeHeights: () => { events.push('geometry'); },
    syncSpacerWidth: () => { events.push('spacer'); },
    padLeft: () => 10,
    invalidateRender: () => { events.push('epoch'); },
    preview: () => { events.push('preview'); },
    scheduleSettle: () => { events.push('settle'); },
    onScaleChange: () => { events.push('change'); },
    relayout: () => { events.push('relayout'); },
    mountVisible: () => { events.push('mount'); },
    refitOnResize: () => true,
  });
  return {
    zoom, host, spacer, events,
    setCount: (n: number) => { count = n; },
    setBase: (n: number) => { base = n; },
    setFitWidth: (n: number) => { fitWidth = n; },
  };
}

it('latches a pre-load scale and anchors a visible zoom before painting', () => {
  const f = fixture();
  f.zoom.setScale(2);
  expect(f.zoom.getScale()).toBe(2);
  expect(f.events).toEqual([]);
  f.setCount(2);
  expect(f.zoom.establishBase()).toBe(true);
  expect(f.events).toEqual(['change']);
  f.events.length = 0;
  f.zoom.pendingAnchor = { x: 100, y: 20 };
  f.zoom.setScale(3);
  expect(f.host.scrollTop).toBe(95);
  expect(f.host.scrollLeft).toBe(75);
  expect(f.spacer.style.height).toBe('930px');
  expect(f.events).toEqual(['epoch', 'geometry', 'spacer', 'preview', 'settle', 'change']);
});

it('preserves the zoom multiplier on a width refit and mounts after a clamp no-op', () => {
  const f = fixture();
  f.setCount(2);
  f.zoom.establishBase();
  f.zoom.setScale(2);
  f.events.length = 0;
  f.setBase(1.5);
  f.setFitWidth(300);
  f.zoom.onResize();
  expect(f.zoom.scale).toBe(3);
  expect(f.zoom.prevBase).toBe(1.5);
  expect(f.events.at(-1)).toBe('mount');
});
