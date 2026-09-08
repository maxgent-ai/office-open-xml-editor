import { describe, expect, it, vi } from 'vitest';

import type { Presentation, ShapeElement } from '@maxgent/ooxml/pptx';

import { createElementRef } from '../../src/adapters/pptx-json-adapter';
import type { Command } from '../../src/domain/command';
import type { ElementRef } from '../../src/domain/mutation';
import { InsertSlideMutation } from '../../src/mutations/insert-slide';
import { UpdateTextMutation } from '../../src/mutations/update-text';
import { PptxEditorViewBindingError } from '../../src/rendering/errors';
import { PptxEditorViewBinding } from '../../src/rendering/pptx-editor-view-binding';
import type { PptxEditorViewHost } from '../../src/rendering/types';
import { PptxEditorSession } from '../../src/session/pptx-editor-session';
import { OFFICECLI_BATCH_SEND_STATUSES } from '../../src/submission/constants';
import type { OfficeCliBatchSendResult } from '../../src/submission/types';
import type { OfficeCliBatch } from '../../src/transport/officecli/types';
import { deck, shape } from '../fixtures/presentation';

describe('PptxEditorViewBinding', () => {
  it('syncs the session presentation into the host on bind and after mutations', async () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const session = createSession(presentation);
    const applies: Array<{
      text: string;
      changedSlideIndexes: readonly number[] | undefined;
    }> = [];
    const host = createHost(async (next, options) => {
      applies.push({
        text: textOf(next.slides[0]),
        changedSlideIndexes: options?.changedSlideIndexes,
      });
    });

    const binding = new PptxEditorViewBinding({ session, host });
    await binding.whenIdle();
    expect(applies).toEqual([{ text: 'before', changedSlideIndexes: undefined }]);

    session.apply(updateTextCommand('edit-1', ref, 'after'));
    await binding.whenIdle();
    expect(applies).toEqual([
      { text: 'before', changedSlideIndexes: undefined },
      { text: 'after', changedSlideIndexes: [0] },
    ]);

    await binding.whenIdle();
    expect(applies).toHaveLength(2);
  });

  it('coalesces rapid mutations to the latest session snapshot', async () => {
    const target = shape('7', 'before');
    const presentation = deck([target]);
    const ref = createElementRef(presentation.slides[0], target, 0);
    const session = createSession(presentation);
    const blocked = deferred<void>();
    const texts: string[] = [];
    const host = createHost(async (next) => {
      const text = textOf(next.slides[0]);
      texts.push(text);
      if (text === 'one') await blocked.promise;
    });

    const binding = new PptxEditorViewBinding({ session, host });
    await binding.whenIdle();

    session.apply(updateTextCommand('edit-1', ref, 'one'));
    session.apply(updateTextCommand('edit-2', ref, 'two'));
    session.apply(updateTextCommand('edit-3', ref, 'three'));

    expect(texts).toEqual(['before', 'one']);
    blocked.resolve(undefined);
    await binding.whenIdle();

    expect(texts).toEqual(['before', 'one', 'three']);

  });

  it('isolates host failures and remains usable for a later sync', async () => {
    const failure = new Error('viewer apply failed');
    const hostApply = vi.fn<PptxEditorViewHost['applyPresentation']>()
      .mockRejectedValueOnce(failure)
      .mockResolvedValue(undefined);
    const onRenderError = vi.fn();
    const binding = new PptxEditorViewBinding({
      session: createSession(deck([shape('7', 'before')])),
      host: { applyPresentation: hostApply },
      onRenderError,
    });

    await binding.whenIdle();
    expect(onRenderError).toHaveBeenCalledWith(failure);

    binding.requestRender();
    await binding.whenIdle();
    expect(hostApply).toHaveBeenCalledTimes(2);
    expect(onRenderError).toHaveBeenCalledTimes(1);
  });

  it('escalates the sync after a host failure to a full presentation apply', async () => {
    const firstTarget = shape('7', 'first');
    const secondTarget = shape('8', 'second');
    const presentation = twoSlideDeck(firstTarget, secondTarget);
    const firstRef = createElementRef(presentation.slides[0], firstTarget, 0);
    const secondRef = createElementRef(presentation.slides[1], secondTarget, 0);
    const session = createSession(presentation);
    let failNext = false;
    const applies: Array<readonly number[] | undefined> = [];
    const host = createHost(async (_next, options) => {
      if (failNext) {
        failNext = false;
        throw new Error('viewer apply failed');
      }
      applies.push(options?.changedSlideIndexes);
    });
    const onRenderError = vi.fn();
    const binding = new PptxEditorViewBinding({ session, host, onRenderError });
    await binding.whenIdle();
    expect(applies).toEqual([undefined]);

    // A failed incremental apply leaves the host's state unknown…
    failNext = true;
    session.apply(updateTextCommand('edit-first', firstRef, 'first-after'));
    await binding.whenIdle();
    expect(onRenderError).toHaveBeenCalledTimes(1);
    expect(applies).toEqual([undefined]);

    // …so the next sync must be a full apply, not just the newly changed slide.
    session.apply(updateTextCommand('edit-second', secondRef, 'second-after'));
    await binding.whenIdle();
    expect(applies).toEqual([undefined, undefined]);

    // Once a full apply succeeded, incremental patches resume.
    session.apply(updateTextCommand('edit-first-again', firstRef, 'first-final'));
    await binding.whenIdle();
    expect(applies).toEqual([undefined, undefined, [0]]);
  });

  it('ignores mutations for slides that are not in the changed set after dispose', async () => {
    const firstTarget = shape('7', 'first');
    const secondTarget = shape('8', 'second');
    const presentation = twoSlideDeck(firstTarget, secondTarget);
    const firstRef = createElementRef(presentation.slides[0], firstTarget, 0);
    const secondRef = createElementRef(presentation.slides[1], secondTarget, 0);
    const session = createSession(presentation);
    const hostApply = vi.fn<PptxEditorViewHost['applyPresentation']>()
      .mockResolvedValue(undefined);
    const binding = new PptxEditorViewBinding({
      session,
      host: { applyPresentation: hostApply },
    });
    await binding.whenIdle();
    expect(hostApply).toHaveBeenCalledTimes(1);

    binding.dispose();
    binding.dispose();
    session.apply(updateTextCommand('edit-first', firstRef, 'first-after'));
    session.apply(updateTextCommand('edit-second', secondRef, 'second-after'));
    await Promise.resolve();
    expect(hostApply).toHaveBeenCalledTimes(1);

    expect(() => binding.requestRender()).toThrowError(
      expect.objectContaining<Partial<PptxEditorViewBindingError>>({
        code: 'viewBinding.disposed',
      }),
    );
  });

  it('can skip the initial sync when the host is already aligned', async () => {
    const hostApply = vi.fn<PptxEditorViewHost['applyPresentation']>()
      .mockResolvedValue(undefined);
    const binding = new PptxEditorViewBinding({
      session: createSession(deck([shape('7', 'before')])),
      host: { applyPresentation: hostApply },
      syncOnBind: false,
    });
    await binding.whenIdle();
    expect(hostApply).not.toHaveBeenCalled();
    binding.dispose();
  });

  it('full-syncs a local slide insertion and its undo without transport', async () => {
    const sendBatch = vi.fn<(
      batch: OfficeCliBatch,
    ) => Promise<OfficeCliBatchSendResult>>()
      .mockResolvedValue({ status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED });
    const session = createSession(deck([shape('7', 'first')]), sendBatch);
    const applies: Array<{
      readonly slideCount: number;
      readonly changedSlideIndexes: readonly number[] | undefined;
    }> = [];
    const binding = new PptxEditorViewBinding({
      session,
      host: createHost(async (presentation, options) => {
        applies.push({
          slideCount: presentation.slides.length,
          changedSlideIndexes: options?.changedSlideIndexes,
        });
      }),
    });
    await binding.whenIdle();

    session.apply({
      id: 'insert-slide',
      mutations: [new InsertSlideMutation({
        index: 1,
        target: { slideId: 'client:slide-2' },
      })],
    });
    await binding.whenIdle();
    expect(applies).toEqual([
      { slideCount: 1, changedSlideIndexes: undefined },
      { slideCount: 2, changedSlideIndexes: undefined },
    ]);

    session.undo();
    await binding.whenIdle();
    expect(applies).toEqual([
      { slideCount: 1, changedSlideIndexes: undefined },
      { slideCount: 2, changedSlideIndexes: undefined },
      { slideCount: 1, changedSlideIndexes: undefined },
    ]);

    expect(sendBatch).not.toHaveBeenCalled();

  });
});

function createSession(
  presentation: Presentation,
  sendBatch = vi.fn<(
    batch: OfficeCliBatch,
  ) => Promise<OfficeCliBatchSendResult>>().mockResolvedValue({
    status: OFFICECLI_BATCH_SEND_STATUSES.CONFIRMED,
  }),
): PptxEditorSession {
  return new PptxEditorSession({
    presentation,
    sendBatch,
    createSaveId: () => 'save-1',
  });
}

function createHost(
  applyPresentation: PptxEditorViewHost['applyPresentation'],
): PptxEditorViewHost {
  return { applyPresentation };
}

function updateTextCommand(id: string, target: ElementRef, value: string): Command {
  return {
    id,
    mutations: [new UpdateTextMutation({ target, value })],
  };
}

function textOf(elementContainer: { readonly elements: readonly unknown[] }): string {
  const element = elementContainer.elements[0] as ShapeElement;
  const run = element.textBody?.paragraphs[0].runs[0];
  return run?.type === 'text' ? run.text : '';
}

function twoSlideDeck(
  firstTarget: ShapeElement,
  secondTarget: ShapeElement,
): Presentation {
  const presentation = deck([firstTarget]);
  const secondSlide = deck([secondTarget]).slides[0];
  return {
    ...presentation,
    slides: [
      presentation.slides[0],
      {
        ...secondSlide,
        index: 1,
        slideNumber: 2,
        partName: 'ppt/slides/slide2.xml',
      },
    ],
  };
}

function deferred<T>(): {
  readonly promise: Promise<T>;
  readonly resolve: (value: T) => void;
} {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((resolvePromise) => {
    resolve = resolvePromise;
  });
  return { promise, resolve };
}
