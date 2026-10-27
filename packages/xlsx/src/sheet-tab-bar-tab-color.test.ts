import { describe, expect, it } from 'vitest';
import { SheetTabBar } from './internal/viewer/sheet-tab-bar.js';
import { makeDocument, type FakeEl } from './viewer-destroy-test-dom.js';

describe('SheetTabBar tab colors', () => {
  it('paints only #RRGGBB colors, as a box-shadow that cannot add declarations (#1957)', () => {
    const doc = makeDocument();
    const bar = new SheetTabBar(doc as unknown as Document, {
      hiddenSheetMode: () => 'show',
      isHidden: () => false,
      selectSheet: () => {},
    });
    bar.build(
      ['Injected', 'Plain'],
      ['#FF0000;background-image:url(http://tabcolor-beacon.invalid/pixel)', '#00ff00'],
    );
    bar.setActive(1);
    const [injected, plain] = bar.tabs as unknown as FakeEl[];
    for (const tab of [injected, plain]) {
      expect(tab.style.cssText).not.toContain('url(');
      expect(tab.style['background-image']).toBe('');
    }
    expect(injected.style.boxShadow).toBe('');
    expect(plain.style.boxShadow).toBe('inset 0 -2px 0 0 #00ff00');
  });
});
