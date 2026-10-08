import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import type { DocxElementContext } from './selection-context.js';
import { FakeDocxEngine, installDom, makeEl } from './scroll-viewer-test-dom.js';
import { DocxViewer } from './viewer.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const PAGE = { widthPt: 600, heightPt: 800 };

function elementContext(elementType: DocxElementContext['elementType'] = 'chart'): DocxElementContext {
  return {
    format: 'docx', kind: 'element', pageIndex: 0, elementIndex: 1, elementType,
    point: { xPt: 0, yPt: 0 },
    bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 50 },
    source: { story: 'body', storyInstance: 'body', path: [2, 0] },
    text: 'Quarterly revenue', seriesCount: 1,
    truncated: false, truncationReasons: [], textCharacters: 17,
    maxTextCharacters: 16_384,
  };
}

async function mount(options: ConstructorParameters<typeof DocxViewer>[1] = {}) {
  installDom();
  const canvas = makeEl('canvas');
  canvas.clientWidth = 900;
  canvas.clientHeight = 1_200;
  const engine = new FakeDocxEngine(2, [PAGE]);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const viewer = new DocxViewer(canvas as unknown as HTMLCanvasElement, options);
  await viewer.load('document.docx');
  return { canvas, engine, viewer, wrapper: canvas.parentElement! };
}

describe('DocxViewer element context', () => {
  it('delivers the original contextmenu event synchronously and resolves the clicked context', async () => {
    let resolveHit!: (context: DocxElementContext | null) => void;
    const received: Array<{ originalEvent: MouseEvent; getContext(): Promise<unknown> }> = [];
    const preventDefault = vi.fn();
    const mounted = await mount({
      enableElementSelection: true,
      onContextMenu(event) {
        received.push(event);
        event.originalEvent.preventDefault();
      },
    });
    mounted.engine.getElementContextAt = vi.fn(() =>
      new Promise<DocxElementContext | null>((resolve) => { resolveHit = resolve; }));
    const originalEvent = {
      button: 2, clientX: 450, clientY: 600, defaultPrevented: false, preventDefault,
    } as unknown as MouseEvent;

    mounted.wrapper.dispatch('contextmenu', originalEvent);

    expect(received).toHaveLength(1);
    expect(received[0].originalEvent).toBe(originalEvent);
    expect(preventDefault).toHaveBeenCalledOnce();
    expect(received[0].getContext()).toBe(received[0].getContext());
    resolveHit(elementContext());
    await expect(received[0].getContext()).resolves.toMatchObject({
      format: 'docx', kind: 'element', elementType: 'chart',
    });
    expect(mounted.viewer.getSelectionContext()).toMatchObject({ kind: 'element' });

    mounted.viewer.destroy();
    expect(mounted.wrapper._listeners.get('contextmenu') ?? []).toHaveLength(0);
  });

  it('rejects contextmenu context lookup failures without also calling onError', async () => {
    const failure = new Error('element lookup failed');
    const onError = vi.fn();
    let received: { getContext(): Promise<unknown> } | undefined;
    const mounted = await mount({
      enableElementSelection: true,
      onError,
      onContextMenu(event) { received = event; },
    });
    mounted.engine.getElementContextAt = vi.fn().mockRejectedValue(failure);

    mounted.wrapper.dispatch('contextmenu', {
      button: 2, clientX: 450, clientY: 600, defaultPrevented: false,
    });

    await expect(received?.getContext()).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    mounted.viewer.destroy();
  });

  it('keeps callback notification separate from element-context activation', async () => {
    const onSelectionContextChange = vi.fn();
    const mounted = await mount({ onSelectionContextChange });
    mounted.engine.elementContext = elementContext();
    expect(mounted.wrapper._listeners.get('contextmenu') ?? []).toHaveLength(0);

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 450, clientY: 600, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls).toEqual([]);
    expect(onSelectionContextChange).not.toHaveBeenCalled();
    expect(mounted.viewer.getSelectionContext()).toBeNull();
    mounted.viewer.destroy();
  });

  it('supports getter-only click context and maps CSS pixels to page points', async () => {
    const currentDate = new Date('2026-08-09T12:00:00Z');
    const mounted = await mount({ enableElementSelection: true, currentDate });
    mounted.engine.elementContext = elementContext();

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 450, clientY: 600, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls).toEqual([{
      pageIndex: 0,
      point: { xPt: 300, yPt: 400 },
      options: { currentDate, maxTextCharacters: 65_536 },
    }]);
    expect(mounted.viewer.getSelectionContext()).toMatchObject({
      format: 'docx', kind: 'element', elementType: 'chart',
      point: { xPt: 0, yPt: 0 },
    });
    const outlineLayer = mounted.wrapper.children.find((child) => child.dataset.overlay === 'element')!;
    expect(outlineLayer.children).toHaveLength(1);
    expect(outlineLayer.children[0].style).toMatchObject({
      left: `${10 / PAGE.widthPt * 100}%`,
      top: `${20 / PAGE.heightPt * 100}%`,
      width: `${100 / PAGE.widthPt * 100}%`,
      height: `${50 / PAGE.heightPt * 100}%`,
      border: '2px solid #1a73e8',
    });

    mounted.engine.elementContext = null;
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 450, clientY: 600, defaultPrevented: false,
    });
    await Promise.resolve();
    expect(outlineLayer.children).toHaveLength(0);
    mounted.viewer.destroy();
  });

  it('caches the hard text maximum and applies the caller budget only at read time', async () => {
    const mounted = await mount({ enableElementSelection: true });
    mounted.engine.elementContext = {
      ...elementContext('shape'),
      text: 'x'.repeat(20_000),
      textCharacters: 20_000,
      maxTextCharacters: 65_536,
    };

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 450, clientY: 600, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls[0]?.options.maxTextCharacters).toBe(65_536);
    expect(mounted.viewer.getSelectionContext({ maxTextCharacters: 65_536 })).toMatchObject({
      textCharacters: 20_000,
      maxTextCharacters: 65_536,
      truncated: false,
    });
    expect(mounted.viewer.getSelectionContext()).toMatchObject({
      textCharacters: 16_384,
      maxTextCharacters: 16_384,
      truncated: true,
    });
    mounted.viewer.destroy();
  });

  it('discards a pending hit after page ABA navigation', async () => {
    const mounted = await mount({ enableElementSelection: true });
    let resolve!: (context: DocxElementContext | null) => void;
    mounted.engine.getElementContextAt = vi.fn(() =>
      new Promise<DocxElementContext | null>((done) => { resolve = done; }));

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    await mounted.viewer.goToPage(1);
    await mounted.viewer.goToPage(0);
    resolve(elementContext());
    await Promise.resolve();

    expect(mounted.viewer.getSelectionContext()).toBeNull();
    mounted.viewer.destroy();
  });
});
