import { expect, it } from 'vitest';
import { ScrollNavigationController } from './scroll-navigation-controller';

it('clamps a last-unit jump and centers an authored point within the scroll extent', () => {
  const host = { clientHeight: 100, clientWidth: 200, scrollTop: 0, scrollLeft: 0 } as HTMLDivElement;
  const spacer = { offsetWidth: 500, style: { width: '500px' } } as HTMLDivElement;
  let mounts = 0;
  const nav = new ScrollNavigationController({
    host: () => host, spacer: () => spacer,
    count: () => 3, established: () => true,
    offset: (unit) => unit * 100, height: () => 100,
    indexAt: (y) => Math.min(2, Math.max(0, Math.floor(y / 100))),
    totalHeight: () => 300, width: () => 300,
    padLeft: () => 10, marginOrigin: () => 0,
    mount: () => { mounts++; },
  });
  nav.scrollToUnit(99);
  expect(host.scrollTop).toBe(200);
  nav.scrollToPoint(1, 150, 50);
  expect([host.scrollTop, host.scrollLeft, mounts]).toEqual([100, 60, 2]);
});
