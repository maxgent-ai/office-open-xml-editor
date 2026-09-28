import { describe, expect, it } from 'vitest';
import { WorksheetPreview } from './worksheet-preview.js';
import type { Row, Worksheet } from '../types.js';

function worksheet(): Worksheet {
  return { rows: [] } as unknown as Worksheet;
}

function row(index: number): Row {
  return { index, height: null, cells: [] };
}

describe('provisional worksheet coverage', () => {
  it('publishes a row-local sheet after the first chunk and waits for later viewports', async () => {
    const progress = new WorksheetPreview([]);
    const sheet = worksheet();
    progress.preview(sheet, null, 70_000, 3);
    const first = progress.ready;
    progress.append([row(1), row(128)]);
    expect(await first).toBe(sheet);
    const later = progress.waitFor(256);
    let settled = false;
    void later.then(() => { settled = true; });
    await Promise.resolve();
    expect(settled).toBe(false);
    progress.append([row(256)]);
    await later;
    expect(sheet.rows.map((item) => item.index)).toEqual([1, 128, 256]);
    progress.finish(sheet);
    await expect(progress.waitFor(70_000)).resolves.toBeUndefined();
  });

  it('holds fallback sheets until completion and rejects waiting viewports on failure', async () => {
    const fallback = new WorksheetPreview([]);
    fallback.preview(null, 'conditional-format', 3, 1);
    fallback.append([row(1)]);
    fallback.finish(worksheet());
    await expect(fallback.ready).resolves.toBeDefined();

    const progress = new WorksheetPreview([]);
    progress.preview(worksheet(), null, 1000, 1);
    progress.append([row(1)]);
    await progress.ready;
    const waiting = progress.waitFor(500);
    const error = new Error('pull failed');
    progress.fail(error);
    await expect(waiting).rejects.toBe(error);
  });

  it('lets a covering-row waiter resolve while later chunks keep arriving', async () => {
    const progress = new WorksheetPreview([]);
    progress.preview(worksheet(), null, 70_000, 1);
    progress.append([row(128)]);
    await progress.ready;
    const covering = progress.waitFor(40);
    const later = progress.waitFor(256);
    await covering;
    progress.append([row(256)]);
    await expect(later).resolves.toBeUndefined();
  });
});
