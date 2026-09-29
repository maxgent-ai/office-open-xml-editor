import type { DestroyableResource, TerminalResourceOwner } from '../internal/canvas-viewer-mechanics';

export interface ScrollLoadHooks<Resource extends DestroyableResource, LoadToken = void> {
  name(): string;
  borrowed(): boolean;
  borrowedMessage(): string;
  destroyed(): boolean;
  owner(): TerminalResourceOwner<Resource>;
  beginLoad?(): LoadToken;
  isCurrentLoad?(token: LoadToken): boolean;
  finishLoad?(token: LoadToken): void;
  acquire(source: string | ArrayBuffer, token: LoadToken): Promise<Resource>;
  beforeReplace(previous: Resource | null): void;
  afterReplace(resource: Resource, token: LoadToken): void;
  mountOpeningWindow(): Promise<void>;
  selectionChanged(): void;
}

/** Atomic self-loaded resource replacement and opening-window lifecycle.
 * Failed acquisition retains the old resource; stale concurrent acquisitions
 * never commit or report through a replaced viewer. */
export class ScrollLoadController<Resource extends DestroyableResource, LoadToken = void> {
  constructor(private readonly hooks: ScrollLoadHooks<Resource, LoadToken>) {}

  async load(source: string | ArrayBuffer): Promise<void> {
    const name = this.hooks.name();
    if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
    if (this.hooks.borrowed()) throw new Error(this.hooks.borrowedMessage());
    const token = this.hooks.beginLoad?.() as LoadToken;
    let selectionInvalidated = false;
    try {
      const resource = await this.hooks.owner().replace(
        () => this.hooks.acquire(source, token),
        (previous) => {
          this.hooks.beforeReplace(previous);
          selectionInvalidated = true;
        },
      );
      if (!resource) return;
      if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
      if (this.hooks.isCurrentLoad && !this.hooks.isCurrentLoad(token)) return;
      this.hooks.afterReplace(resource, token);
      await this.hooks.mountOpeningWindow();
      if (this.hooks.isCurrentLoad && !this.hooks.isCurrentLoad(token)) return;
    } catch (error) {
      if (this.hooks.destroyed()) throw new Error(`${name} is destroyed`);
      if (this.hooks.isCurrentLoad && !this.hooks.isCurrentLoad(token)) return;
      throw error instanceof Error ? error : new Error(String(error));
    } finally {
      this.hooks.finishLoad?.(token);
    }
    // Consumer callbacks run only after the resource and first window commit;
    // their failures are not acquisition or render failures.
    if (selectionInvalidated && !this.hooks.destroyed()) this.hooks.selectionChanged();
  }
}
