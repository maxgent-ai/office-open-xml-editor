import { expect, it } from 'vitest';
import { TerminalResourceOwner } from '../internal/canvas-viewer-mechanics';
import { ScrollLoadController } from './scroll-load-controller';

it('keeps a live resource when a replacement load fails', async () => {
  let retired = 0;
  const old = { destroy: () => { retired++; } };
  const owner = new TerminalResourceOwner('TestViewer', old, true);
  const loader = new ScrollLoadController({
    name: () => 'TestViewer', borrowed: () => false,
    borrowedMessage: () => '', destroyed: () => false,
    owner: () => owner,
    acquire: async () => { throw new Error('invalid package'); },
    beforeReplace: () => { throw new Error('must not commit'); },
    afterReplace: () => {}, mountOpeningWindow: async () => {},
    selectionChanged: () => {},
  });
  await expect(loader.load(new ArrayBuffer(0))).rejects.toThrow('invalid package');
  expect(owner.current).toBe(old);
  expect(retired).toBe(0);
  owner.close();
});
