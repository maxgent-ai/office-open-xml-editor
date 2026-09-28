import { expect, it } from 'vitest';
import { ScrollViewerShell } from './scroll-viewer-shell';

it('owns the scroll surface and removes its listeners on destroy', () => {
  const documentListeners = new Map<string, EventListener>();
  const doc = {
    createElement: () => element(),
    addEventListener: (name: string, listener: EventListener) => documentListeners.set(name, listener),
    removeEventListener: (name: string) => documentListeners.delete(name),
  };
  function element() {
    const listeners = new Map<string, EventListener>();
    return {
      ownerDocument: doc, style: {}, dataset: {}, parentElement: null,
      children: [] as unknown[],
      contains: () => false,
      appendChild(child: unknown) { this.children.push(child); },
      remove() { this.removed = true; },
      removed: false,
      addEventListener(name: string, listener: EventListener) { listeners.set(name, listener); },
      removeEventListener(name: string) { listeners.delete(name); },
      dispatch(name: string) { listeners.get(name)?.({} as Event); },
    };
  }
  const container = element();
  let scrolls = 0;
  let outsideClicks = 0;
  const shell = new ScrollViewerShell(container as unknown as HTMLElement, {
    comments: true,
    onScroll: () => scrolls++,
    onOutsideComment: () => outsideClicks++,
  });
  expect(container.children).toEqual([shell.wrapper]);
  (shell.scrollHost as unknown as ReturnType<typeof element>).dispatch('scroll');
  documentListeners.get('pointerdown')?.({ target: container } as unknown as Event);
  expect([scrolls, outsideClicks]).toEqual([1, 1]);
  shell.destroy();
  (shell.scrollHost as unknown as ReturnType<typeof element>).dispatch('scroll');
  expect([scrolls, outsideClicks, documentListeners.has('pointerdown')]).toEqual([1, 1, false]);
});
