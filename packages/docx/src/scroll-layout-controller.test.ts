import { describe, expect, it } from 'vitest';
import { DocxScrollLayoutController } from './scroll-layout-controller.js';
import { unchangedLeadingPageCount } from './layout/unchanged-pages.js';

describe('progressive publications keep unchanged page canvases', () => {
  const page = (index: number, text: string) => ({ pageIndex: index, layers: { body: [{ text }] } });

  it('counts the leading pages whose records are structurally equal', () => {
    const before = [page(0, 'a'), page(1, 'b'), page(2, 'c')];
    // Fresh objects with equal content, then a changed page, then appended ones.
    const after = [page(0, 'a'), page(1, 'b'), page(2, 'C'), page(3, 'd')];
    expect(unchangedLeadingPageCount(before, after)).toBe(2);
    expect(unchangedLeadingPageCount(before.slice(0, 2), after)).toBe(2);
    expect(unchangedLeadingPageCount(null, after)).toBe(0);
  });

  function controller(presented: number) {
    let invalidations = 0;
    const refreshed: number[] = [];
    const slots = new Map<number, object>([[0, {}], [1, {}]]);
    const layout = new DocxScrollLayoutController({
      current: () => null, destroyed: () => false, report: () => {}, reportBackground: () => {},
      invalidateFind: () => {}, refreshComments: () => {}, adoptView: () => {}, relayout: () => {},
      invalidateRender: () => { invalidations += 1; },
      mounted: () => slots,
      stillMounted: () => true,
      refreshSlot: (index) => { refreshed.push(index); },
    });
    layout.presentedPageCount = presented;
    return { layout, invalidations: () => invalidations, refreshed };
  }

  it('does not discard the presented pages when a publication only appends', () => {
    const view = controller(2);
    view.layout.apply({ pageCount: 6, exact: false, complete: false, unchangedPages: 2 });
    expect(view.invalidations()).toBe(0);
    expect(view.refreshed).toEqual([]);
    expect(view.layout.presentedPageCount).toBe(6);
  });

  it('repaints when a presented page changed or the change is unknown', () => {
    const changed = controller(2);
    changed.layout.apply({ pageCount: 6, exact: false, complete: false, unchangedPages: 1 });
    expect(changed.invalidations()).toBe(1);
    expect(changed.refreshed).toEqual([0, 1]);

    const unknown = controller(2);
    unknown.layout.apply({ pageCount: 6, exact: true, complete: true });
    expect(unknown.invalidations()).toBe(1);
  });
});
