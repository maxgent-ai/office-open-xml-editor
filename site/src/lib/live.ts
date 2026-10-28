// Client-side mounts for the three real viewers used by the home-page showcase.
// Each format is mounted at most once and retained while the page is open, so a
// tab switch is a DOM visibility change rather than another parse.
import { EMU_PER_PX, PT_TO_PX } from '@silurus/ooxml-core';
import { PptxPresentation, PptxScrollViewer } from '@silurus/ooxml-pptx';
import { DocxDocument, DocxScrollViewer } from '@silurus/ooxml-docx';
import { XlsxViewer, googleFonts } from '@silurus/ooxml-xlsx';

export type LiveController = {
  destroy: () => void;
  /** Re-fit a retained viewer after its hidden pane becomes visible again. */
  activate?: () => void;
};

type BorrowedPaginatedViewer =
  | Omit<PptxScrollViewer, 'load'>
  | Omit<DocxScrollViewer, 'load'>;

const MAX_SAMPLE_WIDTH = 880;

function statusLine(text: string): HTMLDivElement {
  const status = document.createElement('div');
  status.className = 'lv-status';
  status.setAttribute('role', 'status');
  status.setAttribute('aria-live', 'polite');
  const circle = document.createElement('span');
  circle.className = 'lv-progress-circle';
  circle.setAttribute('aria-hidden', 'true');
  const label = document.createElement('span');
  label.textContent = text;
  status.append(circle, label);
  return status;
}

function mountPaginated(
  root: HTMLElement,
  kind: 'docx' | 'pptx',
  url: string,
): LiveController {
  root.innerHTML = '';
  const host = document.createElement('div');
  host.className = 'lv-scroll-viewer';
  const status = statusLine('Parsing document…');
  root.append(host, status);

  let destroyed = false;
  let viewer: BorrowedPaginatedViewer | null = null;
  let engine: PptxPresentation | DocxDocument | null = null;
  let maxScale = Number.POSITIVE_INFINITY;
  let resizeObserver: ResizeObserver | null = null;
  let frame: number | null = null;

  const enforceWidthCap = (): void => {
    if (!viewer || destroyed) return;
    const current = viewer.getScale();
    if (current > maxScale) viewer.setScale(maxScale);
  };
  const relayout = (): void => {
    if (!viewer || destroyed) return;
    viewer.relayout();
    if (frame !== null) cancelAnimationFrame(frame);
    frame = requestAnimationFrame(() => {
      frame = null;
      enforceWidthCap();
    });
  };
  const observeWidth = (): void => {
    if (typeof ResizeObserver === 'undefined') return;
    resizeObserver = new ResizeObserver(() => relayout());
    resizeObserver.observe(host);
  };

  if (kind === 'pptx') {
    PptxPresentation.load(url, { fonts: [googleFonts()] }).then((presentation) => {
      if (destroyed) {
        presentation.destroy();
        return;
      }
      engine = presentation;
      maxScale = MAX_SAMPLE_WIDTH / (presentation.slideWidth / EMU_PER_PX);
      viewer = PptxScrollViewer.fromPresentation(host, presentation, {
        gap: 26,
        overscan: presentation.slideCount,
        enableTextSelection: true,
        pageShadow: 'var(--document-shadow)',
        background: 'transparent',
      });
      enforceWidthCap();
      observeWidth();
      status.remove();
    }).catch((error: unknown) => {
      if (!destroyed) status.textContent = message(error);
    });
  } else {
    DocxDocument.load(url, { fonts: [googleFonts()] }).then((document) => {
      if (destroyed) {
        document.destroy();
        return;
      }
      engine = document;
      let widestPage = 0;
      for (let index = 0; index < document.pageCount; index++) {
        widestPage = Math.max(widestPage, document.pageSize(index).widthPt * PT_TO_PX);
      }
      maxScale = widestPage > 0 ? MAX_SAMPLE_WIDTH / widestPage : Number.POSITIVE_INFINITY;
      viewer = DocxScrollViewer.fromDocument(host, document, {
        gap: 26,
        overscan: document.pageCount,
        enableTextSelection: true,
        pageShadow: 'var(--document-shadow)',
        background: 'transparent',
      });
      enforceWidthCap();
      observeWidth();
      status.remove();
    }).catch((error: unknown) => {
      if (!destroyed) status.textContent = message(error);
    });
  }

  return {
    activate: () => relayout(),
    destroy: () => {
      if (destroyed) return;
      destroyed = true;
      resizeObserver?.disconnect();
      if (frame !== null) cancelAnimationFrame(frame);
      viewer?.destroy();
      engine?.destroy();
      root.replaceChildren();
    },
  };
}

export function mountPptx(root: HTMLElement, url: string): LiveController {
  return mountPaginated(root, 'pptx', url);
}

export function mountDocx(root: HTMLElement, url: string): LiveController {
  return mountPaginated(root, 'docx', url);
}

export function mountXlsx(root: HTMLElement, url: string): LiveController {
  root.innerHTML = '';
  const host = document.createElement('div');
  host.className = 'lv-xlsx';
  root.append(host);

  const viewer = new XlsxViewer(host, {
    fonts: [googleFonts()],
    showZoomSlider: true,
    onError: (error: Error) => host.setAttribute('data-error', error.message),
  });
  viewer.load(url).catch((error) => host.setAttribute('data-error', message(error)));

  return {
    destroy: () => {
      viewer.destroy();
      root.replaceChildren();
    },
  };
}

function message(error: unknown): string {
  return `Failed: ${error instanceof Error ? error.message : String(error)}`;
}
