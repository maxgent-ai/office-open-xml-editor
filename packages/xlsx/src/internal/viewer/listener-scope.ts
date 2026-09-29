/**
 * Records every DOM listener a viewer collaborator installs so its `destroy()`
 * can detach them all. Removing the viewer subtree alone would only make
 * element listeners unreachable; detaching them explicitly keeps teardown
 * observable and also covers targets outside the subtree (the owner document).
 */
export class ListenerScope {
  private readonly cleanups: Array<() => void> = [];

  on<K extends keyof HTMLElementEventMap>(
    target: HTMLElement,
    type: K,
    listener: (event: HTMLElementEventMap[K]) => void,
    options?: AddEventListenerOptions | boolean,
  ): () => void {
    target.addEventListener(type, listener as EventListener, options);
    let active = true;
    const cleanup = () => {
      if (!active) return;
      active = false;
      target.removeEventListener(type, listener as EventListener, options);
    };
    this.cleanups.push(cleanup);
    return cleanup;
  }

  /** Detach every listener registered through this scope. Idempotent. */
  dispose(): void {
    for (const cleanup of this.cleanups.splice(0)) cleanup();
  }
}
