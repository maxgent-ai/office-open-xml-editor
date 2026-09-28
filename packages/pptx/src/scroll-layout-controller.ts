import { subscribePptxLayout, type PptxLayoutPublication } from './presentation-layout-events';
import type { PptxPresentation } from './presentation';

/** Tracks progressive slide availability for one presentation at a time. */
export class PptxScrollLayoutController {
  private unsubscribe: (() => void) | null = null;

  constructor(private readonly hooks: {
    current: () => PptxPresentation | null;
    destroyed: () => boolean;
    report: (error: unknown) => void;
    reportBackground: (error: unknown) => void;
    wakeComments: () => void;
    resetComments: () => void;
    failComments: () => void;
    scanComments: (presentation: PptxPresentation) => void;
    renderAvailable: (presentation: PptxPresentation) => void;
    emitVisible: () => void;
  }) {}

  bind(presentation: PptxPresentation): void {
    this.unbind();
    let initial = true;
    this.unsubscribe = subscribePptxLayout(
      presentation,
      () => ({
        availableSlides: presentation.availableSlideCount,
        slideCount: presentation.slideCount,
        exact: presentation.layoutComplete,
        complete: presentation.layoutComplete,
      }),
      (publication) => {
        if (initial) { initial = false; return; }
        this.onPublication(presentation, publication);
      },
      this.hooks.report,
    );
  }

  unbind(): void {
    this.unsubscribe?.();
    this.unsubscribe = null;
    this.hooks.resetComments();
    this.hooks.wakeComments();
  }

  private onPublication(presentation: PptxPresentation, publication: PptxLayoutPublication): void {
    if (this.hooks.destroyed() || presentation !== this.hooks.current()) return;
    this.hooks.wakeComments();
    if (publication.error !== undefined) {
      this.hooks.failComments();
      this.hooks.reportBackground(publication.error);
      return;
    }
    this.hooks.scanComments(presentation);
    this.hooks.renderAvailable(presentation);
    this.hooks.emitVisible();
  }

  destroy(): void { this.unbind(); }
}
