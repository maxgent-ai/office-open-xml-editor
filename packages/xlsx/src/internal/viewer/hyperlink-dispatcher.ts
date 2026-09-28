import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { openExternalHyperlink } from '@silurus/ooxml-core';
import type { DefinedName, Hyperlink, Worksheet } from '../../types.js';
import type { CellAddress } from '../../selection.js';
import { resolveXlsxInternalHyperlink } from '../../internal-hyperlink.js';

/** Options, navigation and error reporting the dispatcher uses from its engine. */
export interface HyperlinkDispatcherHost {
  readonly hostWindow: Window & typeof globalThis;
  /** `false` disables every hyperlink affordance (the viewer option). */
  enabled(): boolean;
  /** Caller override that fully owns hyperlink behaviour, if any. */
  onHyperlinkClick(): ((target: HyperlinkTarget) => void) | undefined;
  currentSheet(): number;
  sheetNames(): string[];
  definedNames(): readonly DefinedName[];
  goToSheet(index: number): Promise<void>;
  scrollToCell(ref: string): Promise<void>;
  reportError(error: unknown): void;
}

/**
 * IX1 cell hyperlinks for the displayed sheet: the `"row:col"` index, the
 * `enableHyperlinks` gate every consumer funnels through, and click dispatch to
 * the caller's `onHyperlinkClick` or the built-in external/internal default.
 */
export class HyperlinkDispatcher {
  /** `"row:col"` → hyperlink for the displayed sheet. Keys mirror the
   *  renderer's `hyperlinkMap` (1-based row/col, the first cell of a hyperlink
   *  `ref` range per the parser), so a `getCellAt` {row,col} looks up directly. */
  private hyperlinkMap = new Map<string, Hyperlink>();

  constructor(private readonly host: HyperlinkDispatcherHost) {}

  /** Index `ws`'s hyperlinks so a clicked/hovered cell resolves in O(1). */
  build(ws: Worksheet): void {
    this.hyperlinkMap = new Map();
    for (const hl of ws.hyperlinks ?? []) {
      this.hyperlinkMap.set(`${hl.row}:${hl.col}`, hl);
    }
  }

  /** The hyperlink at a cell, or null. `getCellAt` returns 1-based {row,col},
   *  matching the parser/renderer keying.
   *
   *  Returns null unconditionally when `enableHyperlinks` is `false`: this is the
   *  single gate that disables hyperlink interactivity. Both consumers — the
   *  pointermove pointer-cursor affordance and the click dispatch
   *  ({@link dispatch}) — funnel through this hit-test, so a null result means
   *  no cursor change, no default navigation, and no `onHyperlinkClick`. */
  at(cell: CellAddress): Hyperlink | null {
    if (!this.host.enabled()) return null;
    return this.hyperlinkMap.get(`${cell.row}:${cell.col}`) ?? null;
  }

  /**
   * Dispatch a click on a hyperlinked cell. Builds a {@link HyperlinkTarget}
   * from the parsed hyperlink (external `url` wins over internal `location`,
   * matching Excel: a `<hyperlink>` carrying both navigates to the external
   * target) and routes it to the caller's `onHyperlinkClick` (which fully owns
   * behaviour) or the built-in default. Returns true when a hyperlink was found
   * and dispatched.
   */
  dispatch(cell: CellAddress): boolean {
    const hl = this.at(cell);
    if (!hl) return false;
    let target: HyperlinkTarget;
    if (hl.url) {
      target = { kind: 'external', url: hl.url };
    } else if (hl.location) {
      target = { kind: 'internal', ref: hl.location };
    } else {
      return false; // parser only emits a hyperlink with url or location
    }
    const custom = this.host.onHyperlinkClick();
    if (custom) {
      custom(target);
      return true;
    }
    // Built-in default. External: open in a new tab, sanitised against the safe
    // scheme allowlist (a blocked scheme like `javascript:` is a no-op, not a
    // navigation). Internal: best-effort sheet navigation, below.
    if (target.kind === 'external') {
      openExternalHyperlink(target.url, undefined, this.host.hostWindow);
    } else {
      void this.navigateInternal(target.ref).catch(
        (error) => this.host.reportError(error),
      );
    }
    return true;
  }

  /**
   * Default handler for an internal `location` target (§18.3.1.47): resolve a
   * direct cell/range or an in-scope defined name (§18.2.5), switch sheets when
   * needed, then scroll the first referenced cell into view.
   */
  async navigateInternal(location: string): Promise<void> {
    const target = resolveXlsxInternalHyperlink(
      location,
      this.host.currentSheet(),
      this.host.sheetNames(),
      this.host.definedNames(),
    );
    if (!target) return;
    if (target.sheetIndex !== this.host.currentSheet()) {
      await this.host.goToSheet(target.sheetIndex);
    }
    await this.host.scrollToCell(target.cellRef);
  }

  destroy(): void {
    this.hyperlinkMap.clear();
  }
}
