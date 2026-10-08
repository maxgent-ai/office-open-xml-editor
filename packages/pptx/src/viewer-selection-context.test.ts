import { afterEach, describe, expect, it, vi } from 'vitest';
import { PptxPresentation } from './presentation.js';
import { PptxViewer } from './viewer.js';
import type { PptxElementContext } from './element-selection.js';
import type { PptxSelectionContext } from './element-selection.js';
import { FakePptxEngine, installDom, makeEl, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_WIDTH = 9_144_000;
const SLIDE_HEIGHT = 6_858_000;

function elementContext(shapeId: string): PptxElementContext {
  return {
    format: 'pptx', kind: 'element', slideIndex: 0, elementIndex: 0,
    origin: 'slide', elementType: 'shape', point: { x: 0, y: 0 },
    bounds: {
      x: 0, y: 0, width: 1_000_000, height: 1_000_000,
      rotation: 0, flipH: false, flipV: false,
    },
    shapeId, geometry: 'rect', truncated: false, truncationReasons: [],
    textCharacters: 0, maxTextCharacters: 16_384,
  };
}

async function mount(mode: 'main' | 'worker' = 'main', slideCount = 1) {
  installDom();
  const canvas = makeEl('canvas');
  canvas.clientWidth = 960;
  canvas.clientHeight = 720;
  const engine = new FakePptxEngine(slideCount, SLIDE_WIDTH, SLIDE_HEIGHT, mode);
  const onSelectionContextChange = vi.fn();
  const onError = vi.fn();
  vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
  const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
    mode,
    enableElementSelection: true,
    onSelectionContextChange,
    onError,
  });
  await viewer.load('deck.pptx');
  return {
    canvas,
    engine,
    onSelectionContextChange,
    onError,
    viewer,
    wrapper: canvas.parentElement as FakeEl,
  };
}

describe('PptxViewer selection context', () => {
  it('delivers originalEvent synchronously while the element context is pending', async () => {
    let resolveHit!: (context: PptxElementContext | null) => void;
    const received: Array<{ originalEvent: MouseEvent; getContext(): Promise<unknown> }> = [];
    installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((resolve) => { resolveHit = resolve; }));
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      onContextMenu(event) { received.push(event); },
    });
    await viewer.load('deck.pptx');
    const wrapper = canvas.parentElement as FakeEl;
    const originalEvent = {
      button: 2, clientX: 480, clientY: 360, defaultPrevented: false,
    } as unknown as MouseEvent;

    wrapper.dispatch('contextmenu', originalEvent);

    expect(received).toHaveLength(1);
    expect(received[0].originalEvent).toBe(originalEvent);
    const contextPromise = received[0].getContext();
    expect(received[0].getContext()).toBe(contextPromise);
    resolveHit(elementContext('context-menu'));
    await expect(contextPromise).resolves.toMatchObject({
      format: 'pptx', kind: 'element', shapeId: 'context-menu',
    });
    viewer.destroy();
    expect(wrapper._listeners.get('contextmenu') ?? []).toHaveLength(0);
  });

  it('does not start a context lookup when the callback never requests it', async () => {
    installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      onContextMenu() {},
    });
    await viewer.load('deck.pptx');

    (canvas.parentElement as FakeEl).dispatch('contextmenu', {
      button: 2, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toEqual([]);
    viewer.destroy();
  });

  it('rejects contextmenu context lookup failures without also calling onError', async () => {
    installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    const failure = new Error('element lookup failed');
    engine.getElementContextAt = vi.fn().mockRejectedValue(failure);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const onError = vi.fn();
    let received: { getContext(): Promise<unknown> } | undefined;
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      onError,
      onContextMenu(event) { received = event; },
    });
    await viewer.load('deck.pptx');

    (canvas.parentElement as FakeEl).dispatch('contextmenu', {
      button: 2, clientX: 480, clientY: 360, defaultPrevented: false,
    });

    await expect(received?.getContext()).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('keeps callback notification separate from element-context activation', async () => {
    installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    engine.elementContext = elementContext('7');
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const onSelectionContextChange = vi.fn();
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      onSelectionContextChange,
    });
    await viewer.load('deck.pptx');

    (canvas.parentElement as FakeEl).dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toEqual([]);
    expect(onSelectionContextChange).not.toHaveBeenCalled();
    expect(viewer.getSelectionContext()).toBeNull();
    viewer.destroy();
  });

  it('supports getter-only element context when explicitly enabled', async () => {
    installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    engine.elementContext = elementContext('7');
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
    });
    await viewer.load('deck.pptx');

    (canvas.parentElement as FakeEl).dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    await Promise.resolve();

    expect(engine.elementContextCalls).toHaveLength(1);
    expect(viewer.getSelectionContext()).toMatchObject({ kind: 'element', shapeId: '7' });
    const outlineLayer = (canvas.parentElement as FakeEl).children.find((child) => child.dataset.overlay === 'element')!;
    expect(outlineLayer.children).toHaveLength(1);
    expect(outlineLayer.children[0].style).toMatchObject({
      left: '0%',
      top: '0%',
      width: `${1_000_000 / SLIDE_WIDTH * 100}%`,
      height: `${1_000_000 / SLIDE_HEIGHT * 100}%`,
      border: '2px solid #1a73e8',
    });
    viewer.destroy();
  });

  it('validates the configurable CSS-pixel line tolerance', () => {
    installDom();
    const canvas = makeEl('canvas');
    expect(() => new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableElementSelection: true,
      elementHitTolerance: Number.NaN,
    })).toThrow(/elementHitTolerance/);
  });

  it.each(['main', 'worker'] as const)(
    'maps a click to slide EMU and emits a detached compact context in %s mode',
    async (mode) => {
      const mounted = await mount(mode);
      mounted.engine.elementContext = elementContext('7');

      mounted.wrapper.dispatch('click', {
        button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
      });
      await Promise.resolve();

      expect(mounted.engine.elementContextCalls).toEqual([{
        slideIndex: 0,
        point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
        options: { tolerance: 6 / 960 * SLIDE_WIDTH, maxTextCharacters: 65_536 },
      }]);
      expect(mounted.onSelectionContextChange).toHaveBeenCalledWith(
        expect.objectContaining({ format: 'pptx', kind: 'element', shapeId: '7' }),
      );
      const snapshot = mounted.viewer.getSelectionContext();
      expect(snapshot).toMatchObject({
        shapeId: '7', point: { x: SLIDE_WIDTH / 2, y: SLIDE_HEIGHT / 2 },
      });
      expect(snapshot).not.toBe(mounted.engine.elementContext);
      mounted.viewer.destroy();
    },
  );

  it('does not let a stale hit-test response overwrite a later click', async () => {
    const mounted = await mount();
    const resolvers: Array<(value: PptxElementContext | null) => void> = [];
    mounted.engine.getElementContextAt = vi.fn(
      (): Promise<PptxElementContext | null> =>
        new Promise((resolve) => resolvers.push(resolve)),
    );

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 200, clientY: 200, defaultPrevented: false,
    });
    resolvers[1](elementContext('new'));
    await Promise.resolve();
    resolvers[0](elementContext('stale'));
    await Promise.resolve();

    expect(mounted.viewer.getSelectionContext()).toMatchObject({ shapeId: 'new' });
    expect(mounted.onSelectionContextChange).toHaveBeenCalledTimes(1);
    mounted.viewer.destroy();
  });

  it('invalidates focus and pending hits across reload and slide ABA navigation', async () => {
    const mounted = await mount('main', 2);
    const resolvers: Array<(value: PptxElementContext | null) => void> = [];
    mounted.engine.getElementContextAt = vi.fn(
      (): Promise<PptxElementContext | null> =>
        new Promise((resolve) => resolvers.push(resolve)),
    );

    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    await mounted.viewer.goToSlide(1);
    await mounted.viewer.goToSlide(0);
    resolvers[0](elementContext('stale-after-aba'));
    await Promise.resolve();
    expect(mounted.viewer.getSelectionContext()).toBeNull();

    mounted.engine.getElementContextAt = vi.fn(async () => elementContext('old-deck'));
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    await Promise.resolve();
    expect(mounted.viewer.getSelectionContext()).toMatchObject({ shapeId: 'old-deck' });

    const replacement = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    vi.mocked(PptxPresentation.load).mockResolvedValueOnce(replacement.asPres());
    const callbackFailure = new Error('consumer callback failed');
    mounted.onSelectionContextChange.mockImplementationOnce(() => { throw callbackFailure; });
    await expect(mounted.viewer.load('replacement.pptx')).rejects.toThrow(callbackFailure);
    expect(mounted.viewer.getSelectionContext()).toBeNull();
    expect(mounted.onSelectionContextChange).toHaveBeenLastCalledWith(null);
    expect(mounted.onError).not.toHaveBeenCalled();
    expect(replacement.renderCalls).toHaveLength(1);
    mounted.viewer.destroy();
  });

  it('completes slide navigation before surfacing a selection callback failure', async () => {
    const mounted = await mount('main', 2);
    mounted.engine.elementContext = elementContext('old-slide');
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });
    await Promise.resolve();
    const callbackFailure = new Error('navigation callback failed');
    mounted.onSelectionContextChange.mockImplementationOnce(() => { throw callbackFailure; });

    await expect(mounted.viewer.goToSlide(1)).rejects.toThrow(callbackFailure);
    expect(mounted.viewer.slideIndex).toBe(1);
    expect(mounted.engine.renderCalls.at(-1)?.slide).toBe(1);
    expect(mounted.onError).not.toHaveBeenCalled();
    mounted.viewer.destroy();
  });

  it('invalidates focus and pending hits when skip mode moves off a hidden slide', async () => {
    const mounted = await mount('main', 3);
    Object.assign(mounted.engine, { isHidden: (index: number) => index === 1 });
    await mounted.viewer.goToSlide(1);
    let resolveHit: (value: PptxElementContext | null) => void = () => undefined;
    mounted.engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((resolve) => { resolveHit = resolve; }));
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });

    await mounted.viewer.setHiddenSlideMode('skip');
    expect(mounted.viewer.slideIndex).not.toBe(1);
    resolveHit(elementContext('stale-hidden-slide'));
    await Promise.resolve();
    expect(mounted.viewer.getSelectionContext()).toBeNull();
    mounted.viewer.destroy();
  });

  it('keeps text as sole focus without a callback and retires a pending element hit', async () => {
    const dom = installDom();
    const canvas = makeEl('canvas');
    canvas.clientWidth = 960;
    canvas.clientHeight = 720;
    const engine = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(engine.asPres());
    const viewer = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      enableTextSelection: true,
      enableElementSelection: true,
    });
    await viewer.load('deck.pptx');
    expect(dom.listenerCount('selectionchange')).toBe(1);

    let resolveHit: (value: PptxElementContext | null) => void = () => undefined;
    engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((resolve) => { resolveHit = resolve; }));
    (canvas.parentElement as FakeEl).dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });

    const originalGetter = viewer.getSelectionContext;
    const text = {
      format: 'pptx', kind: 'text', text: 'selected', slideIndexes: [0], shapeIds: [], runs: [],
      truncated: false, truncationReasons: [], textCharacters: 8,
      maxTextCharacters: 65_536, maxRunLocators: 1_024,
    } satisfies PptxSelectionContext;
    viewer.getSelectionContext = () => text;
    dom.dispatchDocument('selectionchange');
    viewer.getSelectionContext = originalGetter;
    dom.dispatchDocument('selectionchange');
    resolveHit(elementContext('stale-before-text'));
    await Promise.resolve();

    expect(viewer.getSelectionContext()).toBeNull();
    viewer.destroy();
    expect(dom.listenerCount('selectionchange')).toBe(0);
  });

  it('invalidates an old worker hit before reload destroys its bridge', async () => {
    const mounted = await mount('worker');
    let rejectHit: (error: Error) => void = () => undefined;
    mounted.engine.getElementContextAt = vi.fn(() =>
      new Promise<PptxElementContext | null>((_resolve, reject) => { rejectHit = reject; }));
    const originalDestroy = mounted.engine.destroy.bind(mounted.engine);
    mounted.engine.destroy = () => {
      rejectHit(new Error('worker bridge closed'));
      originalDestroy();
    };
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 100, clientY: 100, defaultPrevented: false,
    });

    const replacement = new FakePptxEngine(1, SLIDE_WIDTH, SLIDE_HEIGHT, 'worker');
    vi.mocked(PptxPresentation.load).mockResolvedValueOnce(replacement.asPres());
    await mounted.viewer.load('replacement.pptx');
    await Promise.resolve();

    expect(mounted.onError).not.toHaveBeenCalled();
    expect(mounted.viewer.getSelectionContext()).toBeNull();
    mounted.viewer.destroy();
  });

  it('ignores prevented clicks and closes the query/listener surface on destroy', async () => {
    const mounted = await mount();
    mounted.engine.elementContext = elementContext('7');
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: true,
    });
    await Promise.resolve();
    expect(mounted.engine.elementContextCalls).toEqual([]);

    mounted.viewer.destroy();
    mounted.wrapper.dispatch('click', {
      button: 0, clientX: 480, clientY: 360, defaultPrevented: false,
    });
    expect(mounted.engine.elementContextCalls).toEqual([]);
    expect(() => mounted.viewer.getSelectionContext()).toThrow('PptxViewer is destroyed');
  });
});
