import { expect, it } from 'vitest';
import { ScrollViewportPolicy } from './scroll-viewport-policy';

it('fits authored content inside gutters and honors an explicit width', () => {
  const options = { gap: 12, paddingLeft: 20, paddingRight: 30, width: undefined as number | undefined };
  const viewport = new ScrollViewportPolicy({
    options: () => options,
    container: () => ({ clientWidth: 500 }) as HTMLElement,
    scrollHost: () => ({ clientWidth: 400 }) as HTMLDivElement,
  });
  expect(viewport.fitWidth()).toBe(350);
  expect(viewport.verticalPadding()).toEqual({ leading: 12, trailing: 12 });
  options.width = 300;
  expect(viewport.fitWidth()).toBe(300);
});
