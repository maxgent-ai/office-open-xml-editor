import { expect, it } from 'vitest';
import { VisibleUnitEvents } from './visible-unit-events';

it('publishes a growing or completed layout without a scroll movement', () => {
  const events: Array<[number, number, boolean]> = [];
  const visible = new VisibleUnitEvents((index, total, complete) =>
    events.push([index, total, complete]));
  visible.publish({ topIndex: 0 }, 2, false);
  visible.publish({ topIndex: 0 }, 2, false);
  visible.publish({ topIndex: 0 }, 3, false);
  visible.publish({ topIndex: 0 }, 3, true);
  expect(events).toEqual([[0, 2, false], [0, 3, false], [0, 3, true]]);
});
