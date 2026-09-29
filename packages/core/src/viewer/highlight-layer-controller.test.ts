import { expect, it } from 'vitest';
import { HighlightLayerController } from './highlight-layer-controller';

it('clears mounted search paint and reveals an activated offscreen unit', async () => {
  let active = true;
  let revealed = -1;
  const slot = { highlightLayer: { innerHTML: 'old paint' } as HTMLDivElement };
  const slots = new Map([[0, slot]]);
  const layer = new HighlightLayerController({
    slots: () => slots,
    active: () => active,
    runs: () => [{ text: 'found' }],
    setRuns: () => {},
    paint: (_unit, target) => { target.highlightLayer.innerHTML = 'highlight'; },
    reveal: (unit) => { revealed = unit; },
    unitOf: (location: { page: number }) => location.page,
  });
  layer.redrawAll();
  expect(slot.highlightLayer.innerHTML).toBe('highlight');
  active = false;
  layer.redrawAll();
  expect(slot.highlightLayer.innerHTML).toBe('');
  await layer.activate({ matchIndex: 0, text: 'found', location: { page: 7 } });
  expect(revealed).toBe(7);
});
