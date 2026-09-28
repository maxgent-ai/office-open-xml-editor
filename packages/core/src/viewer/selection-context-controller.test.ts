import { expect, it } from 'vitest';
import { SelectionContextController } from './selection-context-controller';

it('discards a stale element hit after the document changes and removes its listeners', async () => {
  const listeners = new Map<string, EventListener>();
  const documentListeners = new Map<string, EventListener>();
  const ownerDocument = {
    addEventListener: (type: string, listener: EventListener) => { documentListeners.set(type, listener); },
    removeEventListener: (type: string) => { documentListeners.delete(type); },
  };
  const wrapper = {
    ownerDocument,
    contains: (target: Node) => target === marker,
  } as unknown as HTMLDivElement;
  const host = {
    addEventListener: (type: string, listener: EventListener) => { listeners.set(type, listener); },
    removeEventListener: (type: string) => { listeners.delete(type); },
  } as unknown as HTMLDivElement;
  const marker = {} as Node;
  const slot = {
    wrapper,
    canvas: { getBoundingClientRect: () => ({ left: 0, top: 0, width: 100, height: 100 }) },
    elementLayer: null,
  } as unknown as { wrapper: HTMLDivElement; canvas: HTMLCanvasElement; elementLayer: HTMLDivElement | null };
  let resource = { id: 1 };
  let resolveHit!: (context: { kind: 'element'; id: number }) => void;
  const published: unknown[] = [];
  let controller!: SelectionContextController<{ kind: 'element'; id: number }, { kind: 'element'; id: number }, typeof resource, typeof slot>;
  controller = new SelectionContextController({
    wrapper: () => wrapper, scrollHost: () => host, slots: () => new Map([[0, slot]]),
    resource: () => resource, destroyed: () => false,
    textSelectionEnabled: () => true, elementSelectionEnabled: () => true,
    textSelected: () => false, getContext: () => controller.elementContext,
    hitTest: () => new Promise((resolve) => { resolveHit = resolve; }),
    outline: () => null, onChange: (context) => { published.push(context); },
    reportError: () => {},
  });
  controller.bind(true, false);
  expect(documentListeners.has('selectionchange')).toBe(true);
  expect(listeners.has('click')).toBe(true);
  const pending = controller.resolveContextAt({ target: marker, clientX: 20, clientY: 30 } as unknown as MouseEvent);
  resource = { id: 2 };
  resolveHit({ kind: 'element', id: 1 });
  await expect(pending).resolves.toBeNull();
  expect(controller.elementContext).toBeNull();
  expect(published).toEqual([]);
  controller.destroy();
  expect(listeners.size).toBe(0);
  expect(documentListeners.size).toBe(0);
});
