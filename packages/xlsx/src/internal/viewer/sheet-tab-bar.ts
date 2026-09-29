import { HEADER_W } from '../../renderer.js';
import type { HiddenSheetMode } from '../../viewer.js';
import { ListenerScope } from './listener-scope.js';

/** Height of the composite viewer's footer (tab strip + zoom control). */
export const TAB_BAR_H = 30;
// Footer chrome stays in screen pixels: sheet zoom scales grid cells and their
// row/column headers, but must not resize the tab-navigation controls.
const TAB_NAV_W = HEADER_W;
// Gap between adjacent sheet tabs. The first tab also gets this much leading
// space so it is offset from the row-header boundary by the same margin that
// separates tabs from each other.
const TAB_GAP = 1;
/** `'dim'`-mode tab opacity: hidden/veryHidden tabs are greyed but selectable.
 *  A UI-presentation default (ECMA-376 defines no hidden-tab rendering); mirrors
 *  the named pptx `DEFAULT_HIDDEN_DIM` constant. */
const HIDDEN_TAB_DIM_OPACITY = 0.45;

/** Workbook facts and navigation the tab strip needs from its engine. */
export interface SheetTabBarHost {
  hiddenSheetMode(): HiddenSheetMode;
  /** Whether sheet `index` is hidden/veryHidden (`<sheet state>`, §18.2.19). */
  isHidden(index: number): boolean;
  /** Activate sheet `index` (a tab click). */
  selectSheet(index: number): void;
}

/**
 * The composite viewer's footer: Excel-style tab-scroll buttons, the
 * scrollable sheet-tab strip and a slot for the zoom control. Owns its DOM and
 * every listener on it; the active sheet itself belongs to the engine.
 */
export class SheetTabBar {
  readonly tabBar: HTMLDivElement;
  readonly tabStrip: HTMLDivElement;
  /** Direction-aware flex row inside the LTR scroll host. Keeping direction on
   *  this inner row avoids browser-specific negative scrollLeft semantics. */
  readonly tabList: HTMLDivElement;
  readonly navPrev: HTMLButtonElement;
  readonly navNext: HTMLButtonElement;
  tabs: HTMLButtonElement[] = [];
  /** Per-tab colors parallel to `tabs`, from `<sheetPr><tabColor>`. */
  tabColors: (string | null)[] = [];
  private readonly ownerDocument: Document;
  private readonly listeners = new ListenerScope();
  /** Listeners on the current tab buttons; replaced on every rebuild. */
  private tabListeners = new ListenerScope();

  constructor(ownerDocument: Document, private readonly host: SheetTabBarHost) {
    this.ownerDocument = ownerDocument;
    this.tabBar = ownerDocument.createElement('div');
    this.tabBar.style.cssText =
      `display:flex;align-items:flex-end;height:${TAB_BAR_H}px;flex-shrink:0;` +
      `background:var(--ooxml-xlsx-chrome-background,#f0f0f0);` +
      `border-top:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);`;

    // Excel-style scroll buttons. They scroll the tab strip; they do NOT change
    // the active sheet. Disabled (greyed) at the ends / when there is no overflow.
    this.navPrev = this.makeNavButton('◀', 'Scroll tabs left', () => this.scrollTabs(-1));
    this.navNext = this.makeNavButton('▶', 'Scroll tabs right', () => this.scrollTabs(1));
    this.navPrev.dataset.xlsxTabNav = 'prev';
    this.navNext.dataset.xlsxTabNav = 'next';

    // Keep the two-button footer control at the row-header width from the 100%
    // view. It is viewer chrome, so workbook zoom must not resize or shift it.
    const navGroup = ownerDocument.createElement('div');
    navGroup.style.cssText =
      `display:flex;flex-shrink:0;width:${TAB_NAV_W}px;height:100%;`;
    navGroup.appendChild(this.navPrev);
    navGroup.appendChild(this.navNext);

    // The scrollable strip that actually holds the sheet tabs. position:relative
    // so each tab's offsetLeft is measured against the strip's scroll content.
    this.tabStrip = ownerDocument.createElement('div');
    // Keep the scroll host itself LTR so scrollLeft is consistently 0..max in
    // every browser. The inner tabList owns visual LTR/RTL ordering.
    this.tabStrip.style.cssText =
      `position:relative;display:block;flex:1;min-width:0;height:100%;` +
      `margin-left:${TAB_GAP}px;overflow-x:auto;overflow-y:hidden;scrollbar-width:none;`;
    this.tabStrip.classList.add('xlsx-tab-strip');
    this.listeners.on(this.tabStrip, 'scroll', () => this.updateNavButtons());

    // width:max-content preserves overflow scrolling; min-width:100% makes a
    // short RTL tab row fill the strip so row-reverse can right-align it.
    this.tabList = ownerDocument.createElement('div');
    this.tabList.style.cssText =
      `display:flex;align-items:flex-end;height:100%;` +
      `gap:${TAB_GAP}px;box-sizing:border-box;`;
    this.tabList.style.width = 'max-content';
    this.tabList.style.minWidth = '100%';
    this.tabStrip.appendChild(this.tabList);

    this.tabBar.appendChild(navGroup);
    this.tabBar.appendChild(this.tabStrip);
  }

  /** Append trailing footer chrome (the zoom control). */
  append(element: HTMLElement): void {
    this.tabBar.appendChild(element);
  }

  /** Rebuild one tab per sheet, honoring the hidden-sheet mode. */
  build(sheetNames: readonly string[], tabColors: (string | null)[]): void {
    this.tabListeners.dispose();
    this.tabListeners = new ListenerScope();
    this.tabList.innerHTML = '';
    this.tabs = [];
    this.tabColors = tabColors;
    sheetNames.forEach((name, i) => {
      const btn = this.ownerDocument.createElement('button');
      btn.textContent = name;
      btn.title = name;
      btn.style.cssText = this.tabCss(i, false);
      this.tabListeners.on(btn, 'click', () => this.host.selectSheet(i));
      this.tabList.appendChild(btn);
      this.tabs.push(btn);
    });
    this.updateNavButtons();
  }

  private makeNavButton(glyph: string, label: string, onClick: () => void): HTMLButtonElement {
    const btn = this.ownerDocument.createElement('button');
    btn.textContent = glyph;
    btn.setAttribute('aria-label', label);
    btn.title = label;
    btn.classList.add('xlsx-tab-nav');
    btn.style.cssText = this.navButtonStyle(false);
    this.listeners.on(btn, 'click', onClick);
    return btn;
  }

  private navButtonStyle(disabled: boolean): string {
    // Plain triangle icons — no border / tab chrome. The background (incl. the
    // hover tint) lives in the injected `.xlsx-tab-nav` stylesheet so the inline
    // style does not shadow the `:hover` rule.
    const base =
      `flex:1;height:100%;padding:0;` +
      `display:flex;align-items:center;justify-content:center;` +
      `border:none;color:var(--ooxml-xlsx-chrome-text-muted,#666);font-size:9px;line-height:1;` +
      `box-sizing:border-box;outline:none;`;
    return disabled
      ? base + `opacity:0.3;cursor:default;pointer-events:none;`
      : base + `cursor:pointer;`;
  }

  scrollTabs(dir: -1 | 1): void {
    const strip = this.tabStrip;
    const viewLeft = strip.scrollLeft;
    const viewRight = viewLeft + strip.clientWidth;
    let target: number | null = null;
    if (dir === 1) {
      // Nearest tab clipped on the physical right; align its right edge.
      let nearestRight = Number.POSITIVE_INFINITY;
      for (const tab of this.tabs) {
        const right = tab.offsetLeft + tab.offsetWidth;
        if (right > viewRight + 1) nearestRight = Math.min(nearestRight, right);
      }
      if (Number.isFinite(nearestRight)) target = nearestRight - strip.clientWidth;
    } else {
      // Nearest tab clipped on the physical left; align its left edge. Search
      // by geometry, not DOM order, because RTL reverses the visual tab row.
      let nearestLeft = Number.NEGATIVE_INFINITY;
      for (const tab of this.tabs) {
        const left = tab.offsetLeft;
        if (left < viewLeft - 1) nearestLeft = Math.max(nearestLeft, left);
      }
      if (Number.isFinite(nearestLeft)) target = nearestLeft;
    }
    if (target !== null) {
      // Instant (not smooth) so the disabled state is consistent the moment the
      // click resolves — keeps the interaction deterministic to drive/test.
      strip.scrollLeft = Math.max(0, Math.min(target, strip.scrollWidth - strip.clientWidth));
    }
    this.updateNavButtons();
  }

  updateNavButtons(): void {
    const strip = this.tabStrip;
    const atStart = strip.scrollLeft <= 0;
    const atEnd = strip.scrollLeft + strip.clientWidth >= strip.scrollWidth - 1;
    // No overflow => scrollWidth ≈ clientWidth => both ends true => both disabled.
    this.navPrev.style.cssText = this.navButtonStyle(atStart);
    this.navNext.style.cssText = this.navButtonStyle(atEnd);
  }

  /** Restyle every tab for the active sheet and keep that tab in view. */
  setActive(index: number): void {
    this.tabs.forEach((btn, i) => {
      btn.style.cssText = this.tabCss(i, i === index);
    });
    // Keep the active tab visible by scrolling the tab strip HORIZONTALLY only.
    // `scrollIntoView` walks every scrollable ancestor, so it also scrolls the
    // page vertically — on first load that jumped the whole page down to the
    // tab bar (the active sheet is set during load). Adjust the strip's
    // scrollLeft directly so the page never moves.
    // `offsetParent === null` for a `display:none` tab (a hidden sheet reached
    // by an explicit goToSheet in 'skip' mode). Its getBoundingClientRect is all
    // zeros, which would spuriously scroll the strip — skip the scroll for it.
    const tab = this.tabs[index];
    if (tab && tab.offsetParent !== null) {
      const strip = this.tabStrip;
      const tabRect = tab.getBoundingClientRect();
      const stripRect = strip.getBoundingClientRect();
      if (tabRect.left < stripRect.left) {
        strip.scrollLeft -= stripRect.left - tabRect.left;
      } else if (tabRect.right > stripRect.right) {
        strip.scrollLeft += tabRect.right - stripRect.right;
      }
    }
    this.updateNavButtons();
  }

  /** Mirror the workbook footer around the sheet-tab strip for an RTL sheet.
   *  The DOM order remains navigation → tabs → zoom, which is also the
   *  logical reading order; `row-reverse` places that sequence right-to-left.
   *  Move the strip's leading gap with it so the spacing stays symmetric. */
  setDirection(rtl: boolean): void {
    this.tabBar.style.flexDirection = rtl ? 'row-reverse' : 'row';
    this.tabStrip.style.marginLeft = rtl ? '0' : `${TAB_GAP}px`;
    this.tabStrip.style.marginRight = rtl ? `${TAB_GAP}px` : '0';
    this.tabList.style.flexDirection = rtl ? 'row-reverse' : 'row';
  }

  private tabStyle(active: boolean, tabColor?: string | null): string {
    // Active tab renders taller than inactive so the selected sheet draws the
    // eye. Tabs align to flex-end, so shorter inactive tabs sit lower and the
    // active tab sticks up. Font size also bumps a hair on active.
    const activeH = TAB_BAR_H - 2;
    const inactiveH = TAB_BAR_H - 5;
    const base =
      `display:inline-block;flex:none;padding:0 14px;position:relative;` +
      `border:1px solid var(--ooxml-xlsx-chrome-border,#c8ccd0);border-bottom:none;` +
      `border-radius:3px 3px 0 0;` +
      `cursor:pointer;white-space:nowrap;max-width:160px;overflow:hidden;text-overflow:ellipsis;` +
      `outline:none;box-sizing:border-box;`;
    // `<sheetPr><tabColor>` renders as a color bar along the tab's bottom edge
    // (Excel's "sheet tab color" treatment), drawn as an inset bottom shadow so
    // it doesn't fight the tab's own border/background. The active tab keeps a
    // thinner bar since its bottom merges into the white sheet body.
    const bar = tabColor
      ? `box-shadow:inset 0 -${active ? 2 : 3}px 0 0 ${tabColor};`
      : '';
    return active
      ? base +
        `height:${activeH}px;font-size:13px;` +
        `background:var(--ooxml-xlsx-chrome-surface,#fff);` +
        `color:var(--ooxml-xlsx-chrome-text,#000);` +
        `border-bottom:1px solid var(--ooxml-xlsx-chrome-surface,#fff);` +
        `font-weight:600;top:1px;` +
        bar
      : base +
        `height:${inactiveH}px;font-size:11px;` +
        `background:var(--ooxml-xlsx-chrome-surface-muted,#e0e0e0);` +
        `color:var(--ooxml-xlsx-chrome-text-muted,#555);` +
        bar;
  }

  /**
   * Full inline style for the tab of sheet `i`, honoring the hidden-sheet mode:
   * `'skip'` hides the tab of a hidden/veryHidden sheet (`display:none`); `'dim'`
   * greys it but leaves it clickable; `'show'` styles every tab normally. Used
   * by both build and setActive so navigation never wipes the styling.
   */
  private tabCss(i: number, active: boolean): string {
    let css = this.tabStyle(active, this.tabColors[i]);
    const mode = this.host.hiddenSheetMode();
    if (mode !== 'show' && this.host.isHidden(i)) {
      css += mode === 'skip' ? 'display:none;' : `opacity:${HIDDEN_TAB_DIM_OPACITY};`;
    }
    return css;
  }

  /** Detach every tab-bar listener (the DOM leaves with the viewer subtree). */
  destroy(): void {
    this.tabListeners.dispose();
    this.listeners.dispose();
  }
}
