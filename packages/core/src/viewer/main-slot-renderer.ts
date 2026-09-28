import { StaticCanvasRenderDispatcher } from '../internal/canvas-viewer-mechanics';
import type { BitmapSlot } from './bitmap-slot-renderer';

export interface MainSlot extends BitmapSlot { wrapper: HTMLDivElement; }

export interface MainSlotHooks<Slot extends MainSlot, Run> {
  slots(): ReadonlyMap<number, Slot>;
  epoch(): number;
  scale(): number;
  token(slot: Slot): number;
  nextToken(slot: Slot): number;
  slotIndex(slot: Slot): number;
  wantRuns(slot: Slot): boolean;
  render(unit: number, canvas: HTMLCanvasElement, width: number, dpr: number,
    onRun: ((run: Run) => void) | undefined, settled: boolean): Promise<void>;
  commitRuns(unit: number, slot: Slot, runs: Run[], canvas: HTMLCanvasElement,
    width: number, wantedRuns: boolean, settled: boolean): void;
  shadow(): string | false;
  reportError(error: unknown): void;
}

/** Direct main-thread painting and blank-free spare-canvas settle. Direct
 * paint skips overlay side effects when a zoom/recycle moves the epoch or slot
 * identity. During settle, the renderer may clear a canvas backing store before
 * its first await, so a spare stays off DOM until it can replace the stretched
 * preview in one operation. */
export class MainSlotRenderer<Slot extends MainSlot, Run> {
  constructor(private readonly hooks: MainSlotHooks<Slot, Run>) {}

  render(unit: number, slot: Slot, width: number, dpr: number,
    token: number, dispatcher: StaticCanvasRenderDispatcher,
    generation: number, reportErrors: boolean): Promise<void> {
    const epoch = this.hooks.epoch();
    const scale = this.hooks.scale();
    const canvas = slot.canvas;
    const runs: Run[] = [];
    const wantedRuns = this.hooks.wantRuns(slot);
    let render: Promise<void>;
    try {
      render = this.hooks.render(unit, canvas, width, dpr,
        wantedRuns ? (run) => runs.push(run) : undefined, false);
    } catch (error) {
      if (reportErrors) { this.hooks.reportError(error); return Promise.resolve(); }
      return Promise.reject(error);
    }
    return render.then(() => {
      if (!this.isCurrent(unit, slot, token, dispatcher, generation, canvas, epoch)) return;
      slot.renderedScale = scale;
      this.hooks.commitRuns(unit, slot, runs, canvas, width, wantedRuns, false);
    }).catch((error: unknown) => {
      if (!this.isCurrent(unit, slot, token, dispatcher, generation, canvas, epoch)) return;
      if (reportErrors) this.hooks.reportError(error);
      else throw error;
    });
  }

  settle(unit: number, slot: Slot, width: number, dpr: number): void {
    const scale = this.hooks.scale();
    const epoch = this.hooks.epoch();
    const spare = document.createElement('canvas');
    spare.style.cssText = 'display:block;background:#fff;';
    const shadow = this.hooks.shadow();
    if (shadow !== false) spare.style.boxShadow = shadow;
    const token = this.hooks.nextToken(slot);
    const dispatcher = new StaticCanvasRenderDispatcher(spare, false);
    const generation = dispatcher.begin();
    const runs: Run[] = [];
    const wantedRuns = this.hooks.wantRuns(slot);
    let render: Promise<void>;
    try {
      render = this.hooks.render(unit, spare, width, dpr,
        wantedRuns ? (run) => runs.push(run) : undefined, true);
    } catch (error) {
      if (this.isCurrent(unit, slot, token, dispatcher, generation, spare, epoch, true)) {
        this.hooks.reportError(error);
      }
      dispatcher.destroy();
      return;
    }
    void render.then(() => {
      if (!this.isCurrent(unit, slot, token, dispatcher, generation, spare, epoch, true)) {
        dispatcher.destroy();
        return;
      }
      const old = slot.canvas;
      slot.dispatcher.destroy();
      slot.wrapper.insertBefore(spare, old);
      old.remove();
      slot.canvas = spare;
      slot.dispatcher = dispatcher;
      slot.renderedScale = scale;
      this.hooks.commitRuns(unit, slot, runs, spare, width, wantedRuns, true);
    }).catch((error: unknown) => {
      if (this.isCurrent(unit, slot, token, dispatcher, generation, spare, epoch, true)) {
        this.hooks.reportError(error);
      }
      dispatcher.destroy();
    });
  }

  private isCurrent(unit: number, slot: Slot, token: number,
    dispatcher: StaticCanvasRenderDispatcher, generation: number,
    canvas: HTMLCanvasElement, epoch: number, spare = false): boolean {
    return token === this.hooks.token(slot) && dispatcher.isCurrent(generation) &&
      (spare || canvas === slot.canvas) && epoch === this.hooks.epoch() &&
      this.hooks.slots().get(unit) === slot && this.hooks.slotIndex(slot) === unit;
  }
}
