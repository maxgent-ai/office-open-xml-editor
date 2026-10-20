import { expect, it, vi } from 'vitest';
import { OoxmlDecodedImageLimitError } from './pixel-budget.js';
import { decodeRasterOrMetafile } from './raster-or-metafile.js';
import { deserializeWorkerError, serializeWorkerError } from '../worker/error-wire.js';

vi.mock('./emf.js', () => ({
  renderEmfToBitmap: async () => {
    throw new OoxmlDecodedImageLimitError('metafile-path-commands', 4, 5);
  },
}));

it('never applies incomplete-metafile compatibility to a resource failure', async () => {
  // isEmf reads the ENHMETAHEADER record type and EMF signature.
  const header = new Uint8Array(88);
  const view = new DataView(header.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(40, 0x464d4520, true);
  for (const incompleteMetafile of [undefined, 'draw-supported', 'reject'] as const) {
    await expect(decodeRasterOrMetafile(new Blob([header]), { incompleteMetafile }))
      .rejects.toMatchObject({
        code: 'ooxml-decoded-image-limit', metric: 'metafile-path-commands', limit: 4, observed: 5,
      });
  }
});

it('preserves the path-budget failure through the worker error protocol', () => {
  const error = new OoxmlDecodedImageLimitError('metafile-path-commands', 4, 5);
  expect(deserializeWorkerError(serializeWorkerError(error))).toMatchObject({
    code: error.code, metric: error.metric, limit: 4, observed: 5,
  });
});
