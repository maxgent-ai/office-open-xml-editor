import { afterEach, describe, expect, it, vi } from 'vitest';
import { SlotScroller, type SlotWindow } from './slot-scroller';

interface TestSlot { renderedScale: number; id: number; }

afterEach(() => vi.useRealTimers());

function makeScroller() {
  let window: SlotWindow = { start: 0, end: 1, topIndex: 0, totalHeight: 500 };
  let nextId = 0;
  let scale = 1;
  const render = vi.fn((index: number, slot: TestSlot) => { slot.renderedScale = scale; });
  const preview = vi.fn();
  const settle = vi.fn((index: number, slot: TestSlot) => { slot.renderedScale = scale; });
  const reset = vi.fn();
  const spacer = { style: { height: '' } } as HTMLDivElement;
  const scroller = new SlotScroller<TestSlot, SlotWindow>({
    spacer: () => spacer,
    count: () => 10,
    range: () => window,
    createSlot: () => ({ renderedScale: -1, id: nextId++ }),
    attachSlot: () => {},
    resetSlot: reset,
    positionSlot: () => {},
    renderSlot: render,
    previewSlot: preview,
    settleSlot: settle,
    renderedScale: (slot) => slot.renderedScale,
    scale: () => scale,
    syncSpacerWidth: () => {},
    onRange: () => {},
  });
  return {
    scroller, render, preview, settle, reset, spacer,
    setWindow: (next: SlotWindow) => { window = next; },
    setScale: (next: number) => { scale = next; },
    creations: () => nextId,
  };
}

describe('SlotScroller', () => {
  it('bounds mounted slots and reuses a detached slot on distant scroll', () => {
    const fixture = makeScroller();
    fixture.scroller.syncSpacer();
    fixture.scroller.mount();
    expect(fixture.spacer.style.height).toBe('500px');
    expect([...fixture.scroller.slots.keys()]).toEqual([0, 1]);
    const first = fixture.scroller.slots.get(0);

    fixture.setWindow({ start: 2, end: 3, topIndex: 2, totalHeight: 500 });
    fixture.scroller.mount();
    expect([...fixture.scroller.slots.keys()]).toEqual([2, 3]);
    expect(fixture.creations()).toBe(2);
    expect([...fixture.scroller.slots.values()]).toContain(first);
    expect(fixture.reset).toHaveBeenCalledTimes(2);
  });

  it('previews existing slots, renders entrants, and settles once after a zoom burst', () => {
    vi.useFakeTimers();
    const fixture = makeScroller();
    fixture.scroller.mount();
    fixture.setScale(2);
    fixture.setWindow({ start: 1, end: 2, topIndex: 1, totalHeight: 900 });
    fixture.scroller.preview();
    expect(fixture.preview).toHaveBeenCalledTimes(1);
    expect(fixture.render).toHaveBeenCalledTimes(3);

    fixture.scroller.renderEpoch++;
    fixture.scroller.scheduleSettle(150);
    vi.advanceTimersByTime(100);
    fixture.scroller.scheduleSettle(150);
    vi.advanceTimersByTime(149);
    expect(fixture.settle).not.toHaveBeenCalled();
    vi.advanceTimersByTime(1);
    expect(fixture.settle).toHaveBeenCalledTimes(1);
    expect(fixture.settle.mock.calls[0][0]).toBe(1);
  });

  it('cancels a pending settle and clears the pool on destroy', () => {
    vi.useFakeTimers();
    const fixture = makeScroller();
    fixture.scroller.mount();
    fixture.scroller.scheduleSettle(150);
    fixture.scroller.destroy();
    vi.advanceTimersByTime(150);
    expect(fixture.settle).not.toHaveBeenCalled();
    expect(fixture.scroller.slots.size).toBe(0);
    expect(fixture.scroller.free).toHaveLength(0);
    expect(fixture.reset).toHaveBeenCalledTimes(2);
  });
});
