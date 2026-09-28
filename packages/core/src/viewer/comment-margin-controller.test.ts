import { expect, it } from 'vitest';
import { CommentMarginController } from './comment-margin-controller';

it('adds a late left review rail without moving authored content on screen', () => {
  let hasComments = false;
  const container = {
    dir: '', style: { direction: '' },
    ownerDocument: { defaultView: { getComputedStyle: () => ({ direction: 'ltr' }) } },
  } as unknown as HTMLElement;
  const scrollHost = { style: {} as Record<string, string>, scrollLeft: 0 } as unknown as HTMLDivElement;
  const spacer = { style: { width: '' } } as HTMLDivElement;
  const margin = { style: {} as Record<string, string>, dataset: {} as Record<string, string> } as unknown as HTMLDivElement;
  const rail = new CommentMarginController({
    container: () => container, scrollHost: () => scrollHost, spacer: () => spacer,
    enabled: () => true, cards: () => true, hasDisplayableComments: () => hasComments,
    requestedSide: () => 'left', zoom: () => 2, gapPx: 8, widthPx: 100, fontSizePx: 12,
  });

  rail.syncMargin(margin);
  rail.syncSpacerWidth(500, 10, 10);
  expect(margin.style.display).toBe('none');
  expect(spacer.style.width).toBe('520px');
  hasComments = true;
  rail.syncMargin(margin);
  rail.syncSpacerWidth(500, 10, 10);
  expect(margin.style.display).toBe('');
  expect(margin.style.right).toBe('calc(100% + 16px)');
  expect(spacer.style.width).toBe('736px');
  expect(scrollHost.scrollLeft).toBe(216);
  expect(scrollHost.style.getPropertyValue?.('--ooxml-review-origin-x') ??
    (scrollHost.style as CSSStyleDeclaration & Record<string, string>)['--ooxml-review-origin-x']).toBe('216px');

  hasComments = false;
  rail.syncSpacerWidth(500, 10, 10);
  expect(scrollHost.scrollLeft).toBe(0);
  rail.destroy();
});
