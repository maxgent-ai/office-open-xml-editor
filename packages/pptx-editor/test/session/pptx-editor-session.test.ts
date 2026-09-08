import { describe, expect, it, vi } from 'vitest';
import { PptxEditorSession, UpdateTextMutation, createElementRef } from '../../src';
import { deck, plainShape } from '../fixtures/presentation';

describe('manual save session', () => {
  it('applies edits without transport and sends only on save', async () => {
    const presentation = deck([plainShape('7', 'before')]);
    const target = createElementRef(presentation.slides[0], presentation.slides[0].elements[0], 0);
    const sendBatch = vi.fn().mockResolvedValue({ status: 'confirmed' });
    const session = new PptxEditorSession({ presentation, sendBatch, createSaveId: () => 'save-1' });
    session.apply({ id: 'edit', mutations: [new UpdateTextMutation({ target, value: 'after' })] });
    expect(sendBatch).not.toHaveBeenCalled();
    expect(session.getSnapshot()).toMatchObject({ dirty: true, canUndo: true, saveStatus: 'idle' });
    await expect(session.save()).resolves.toEqual({ commandId: 'save-1', status: 'confirmed' });
    expect(sendBatch).toHaveBeenCalledTimes(1);
    expect(session.getSnapshot().dirty).toBe(false);
  });
});

import type { Presentation, ShapeElement } from '@maxgent/ooxml/pptx';
import {
  AddElementMutation, InsertSlideMutation, RemoveElementMutation, RemoveSlideMutation,
  UpdateShapeMutation, Mutation, MUTATION_TYPES,
  type ElementRef, type OfficeCliBatch, type OfficeCliBatchSendResult,
} from '../../src';

function setup(presentation = deck([plainShape('7', 'before'), plainShape('8', 'other')])) {
  const target = createElementRef(presentation.slides[0], presentation.slides[0].elements[0], 0);
  const sendBatch = vi.fn<(batch: OfficeCliBatch) => Promise<OfficeCliBatchSendResult>>()
    .mockResolvedValue({ status: 'confirmed' });
  let next = 0;
  const session = new PptxEditorSession({ presentation, sendBatch, createSaveId: () => `save-${++next}` });
  const edit = (value: string, ref = target) => session.apply({
    id: `edit-${value}`, mutations: [new UpdateTextMutation({ target: ref, value })],
  });
  return { session, sendBatch, edit, target, presentation };
}

function textOf(presentation: Presentation): string[] {
  return presentation.slides.flatMap((slide) => slide.elements.map((element) => {
    const run = (element as ShapeElement).textBody?.paragraphs[0].runs[0];
    return run?.type === 'text' ? run.text : '';
  }));
}

it('keeps text and style operations ordered in one save', async () => {
  const { session, sendBatch, edit, target } = setup();
  edit('one');
  session.apply({ id: 'style', mutations: [new UpdateTextMutation({ target, style: { bold: false } })] });
  edit('two');
  expect(textOf(session.getSnapshot().presentation)[0]).toBe('two');
  expect(sendBatch).not.toHaveBeenCalled();
  await session.save();
  expect(sendBatch).toHaveBeenCalledTimes(1);
  expect(sendBatch.mock.calls[0][0].commands).toMatchObject([
    { props: { text: 'one' } }, { props: { bold: 'false' } }, { props: { text: 'two' } },
  ]);
});

it('uses the updated document for add, edit, move, delete and subsequent edits', async () => {
  const { session, sendBatch, target, edit } = setup();
  const added = { ...target, elementId: '9' };
  session.apply({ id: 'add', mutations: [new AddElementMutation({
    target: added, element: plainShape('9', 'new'), presentationElementIndex: 1,
  })] });
  edit('added text', added);
  session.apply({ id: 'move', mutations: [new UpdateShapeMutation({ target: added, value: { x: 123 } })] });
  session.apply({ id: 'delete', mutations: [new RemoveElementMutation({ target })] });
  edit('remaining', { ...target, elementId: '8' });
  await session.save();
  expect(sendBatch.mock.calls[0][0].commands).toMatchObject([
    { command: 'add', parent: '/slide[1]', props: { id: '9', zorder: '2' } },
    { command: 'set', path: '/slide[1]/shape[@id=9]', props: { text: 'added text' } },
    { command: 'set', path: '/slide[1]/shape[@id=9]', props: { x: '123emu' } },
    { command: 'remove', path: '/slide[1]/shape[@id=7]' },
    { command: 'set', path: '/slide[1]/shape[@id=8]', props: { text: 'remaining' } },
  ]);
  expect(textOf(session.getSnapshot().presentation)).toEqual(['added text', 'remaining']);
});

it('resolves slide indexes after insertion and removal', async () => {
  const { session, sendBatch, target, edit } = setup();
  session.apply({ id: 'insert', mutations: [new InsertSlideMutation({ target: { slideId: 'local:new' }, index: 0 })] });
  edit('on second');
  const added: ElementRef = { slideId: 'local:new', origin: 'slide', elementId: '9' };
  session.apply({ id: 'add', mutations: [new AddElementMutation({ target: added, element: plainShape('9', 'new'), presentationElementIndex: 0 })] });
  session.apply({ id: 'remove', mutations: [new RemoveSlideMutation({ target })] });
  edit('on first', added);
  await session.save();
  expect(sendBatch.mock.calls[0][0].commands).toMatchObject([
    { command: 'add', type: 'slide', index: 0 },
    { command: 'set', path: '/slide[2]/shape[@id=7]' },
    { command: 'add', parent: '/slide[1]' },
    { command: 'remove', path: '/slide[2]' },
    { command: 'set', path: '/slide[1]/shape[@id=9]' },
  ]);
  expect(textOf(session.getSnapshot().presentation)).toEqual(['on first']);
});

it('omits undone edits, preserves redo, and creates a new save after saved undo', async () => {
  const { session, edit, sendBatch } = setup();
  edit('one'); edit('two'); edit('three');
  session.undo(); session.undo(); session.redo();
  await session.save();
  expect(sendBatch.mock.calls[0][0].commands).toMatchObject([{ props: { text: 'one' } }, { props: { text: 'two' } }]);
  expect(session.getSnapshot()).toMatchObject({ dirty: false, canRedo: true });
  session.undo();
  expect(textOf(session.getSnapshot().presentation)[0]).toBe('one');
  expect(session.getSnapshot().dirty).toBe(true);
  await session.save();
  expect(sendBatch.mock.calls[1][0].commands).toMatchObject([{ props: { text: 'one' } }]);
  session.redo();
  await session.save();
  expect(sendBatch.mock.calls[2][0].commands).toMatchObject([{ props: { text: 'two' } }]);
});

it('replays inverses before edits on a new branch from saved history', async () => {
  const { session, edit, sendBatch } = setup();
  edit('one'); edit('two'); await session.save();
  session.undo(); session.undo(); edit('branch');
  expect(session.getSnapshot().canRedo).toBe(false);
  await session.save();
  expect(sendBatch.mock.calls[1][0].commands).toMatchObject([
    { props: { text: 'one' } }, { props: { text: 'before' } }, { props: { text: 'branch' } },
  ]);
});

it('does not send when undo returns exactly to the saved position', async () => {
  const { session, edit, sendBatch } = setup();
  edit('one'); session.undo();
  expect(session.getSnapshot().dirty).toBe(false);
  await expect(session.save()).resolves.toEqual({ status: 'unchanged' });
  expect(sendBatch).not.toHaveBeenCalled();
});

it('retains draft and history after definite rejection; retry requires a new explicit save', async () => {
  const { session, edit, sendBatch } = setup();
  const cause = new Error('not applied');
  sendBatch.mockResolvedValueOnce({ status: 'rejected', cause });
  edit('draft');
  await expect(session.save()).resolves.toMatchObject({ status: 'rejected', cause });
  expect(textOf(session.getSnapshot().presentation)[0]).toBe('draft');
  expect(session.getSnapshot()).toMatchObject({ dirty: true, canUndo: true, saveStatus: 'failed', saveError: cause });
  expect(sendBatch).toHaveBeenCalledTimes(1);
  await session.save();
  expect(sendBatch.mock.calls.map(([batch]) => batch.commandId)).toEqual(['save-1', 'save-2']);
  expect(session.getSnapshot().dirty).toBe(false);
});

it.each(['unknown', 'throw', 'invalid'])('locks an ambiguous %s outcome without automatic replay', async (outcome) => {
  const { session, edit, sendBatch, presentation } = setup();
  if (outcome === 'throw') sendBatch.mockRejectedValueOnce(new Error('timeout'));
  else sendBatch.mockResolvedValueOnce({ status: outcome, cause: 'timeout' } as OfficeCliBatchSendResult);
  edit('draft');
  await expect(session.save()).resolves.toMatchObject({ status: 'unknown' });
  expect(session.getSnapshot()).toMatchObject({ dirty: true, canEdit: false, canUndo: false, canRedo: false, saveStatus: 'unknown' });
  expect(() => edit('blocked')).toThrow();
  expect(() => session.undo()).toThrow();
  expect(() => session.redo()).toThrow();
  await expect(session.save()).rejects.toThrow();
  expect(sendBatch).toHaveBeenCalledTimes(1);
  expect(() => session.resolveUnknown('wrong-id', { status: 'confirmed' })).toThrow();
  session.resync(presentation);
  expect(session.getSnapshot()).toMatchObject({ dirty: false, saveStatus: 'idle', undoDepth: 0 });
});

it.each(['confirmed', 'rejected'] as const)('allows externally verified %s resolution of an unknown save', async (status) => {
  const { session, edit, sendBatch } = setup();
  sendBatch.mockResolvedValueOnce({ status: 'unknown', cause: 'timeout' });
  edit('draft'); await session.save();
  session.resolveUnknown('save-1', status === 'confirmed' ? { status } : { status, cause: 'verified not applied' });
  expect(session.getSnapshot()).toMatchObject({ dirty: status !== 'confirmed', canEdit: true });
  expect(sendBatch).toHaveBeenCalledTimes(1);
});

it('locks concurrent operations and publishes saving before invoking transport', async () => {
  const { session, edit, sendBatch, presentation } = setup();
  let resolve!: (value: OfficeCliBatchSendResult) => void;
  sendBatch.mockImplementation(() => new Promise((done) => { resolve = done; }));
  edit('draft');
  const changes: string[] = [];
  session.subscribe(({ snapshot }) => { changes.push(snapshot.saveStatus); });
  const saving = session.save();
  expect(session.getSnapshot()).toMatchObject({ dirty: true, saveStatus: 'saving', canEdit: false });
  expect(changes).toEqual(['saving']);
  expect(() => edit('blocked')).toThrow();
  expect(() => session.undo()).toThrow();
  expect(() => session.redo()).toThrow();
  expect(() => session.resync(presentation)).toThrow();
  await expect(session.save()).rejects.toThrow();
  resolve({ status: 'confirmed' }); await saving;
  expect(changes).toEqual(['saving', 'idle']);
  expect(sendBatch).toHaveBeenCalledTimes(1);
});

it('supports local undo of noninvertible deletion, but stops saved undo at its boundary', async () => {
  const { session, target, edit } = setup();
  session.apply({ id: 'remove', mutations: [new RemoveElementMutation({ target })] });
  session.undo();
  expect(textOf(session.getSnapshot().presentation)).toEqual(['before', 'other']);
  session.redo(); await session.save();
  expect(session.getSnapshot().canUndo).toBe(false);
  edit('later', { ...target, elementId: '8' }); await session.save();
  session.undo();
  expect(session.getSnapshot().canUndo).toBe(false);
  await session.save();
});

it('supports local undo/redo of deleting a populated slide', async () => {
  const { session, target, presentation } = setup();
  session.apply({ id: 'delete', mutations: [new RemoveSlideMutation({ target })] });
  session.undo();
  expect(session.getSnapshot().presentation).toBe(presentation);
  session.redo(); await session.save();
  expect(session.getSnapshot().canUndo).toBe(false);
});

it('shares the save id and monotonically increasing mutation indexes across edits', async () => {
  const { session } = setup();
  const contexts: unknown[] = [];
  class Probe extends Mutation {
    readonly type = MUTATION_TYPES.UPDATE_SHAPE;
    readonly target = { slideId: 'ppt/slides/slide1.xml' };
    apply(presentation: Presentation) { return { presentation, changedSlideIds: [], changedElements: [] }; }
    inverse() { return this; }
    toOfficeCli(_presentation: Presentation, context: unknown) {
      contexts.push(context);
      return { command: 'set' as const, path: '/slide[1]/shape[@id=7]', props: { x: '0emu' } };
    }
  }
  session.apply({ id: 'a', mutations: [new Probe()] });
  session.apply({ id: 'b', mutations: [new Probe(), new Probe()] });
  contexts.length = 0;
  await session.save();
  expect(contexts).toEqual([0, 1, 2].map((mutationIndex) => ({ commandId: 'save-1', mutationIndex })));
});

it('isolates listener errors and settles an in-flight save after disposal', async () => {
  const presentation = deck([plainShape('7', 'before')]);
  let resolve!: (value: OfficeCliBatchSendResult) => void;
  const onListenerError = vi.fn();
  const session = new PptxEditorSession({ presentation, createSaveId: () => 'save',
    sendBatch: () => new Promise((done) => { resolve = done; }), onListenerError });
  const listener = vi.fn(() => { throw new Error('view'); });
  session.subscribe(listener);
  session.apply({ id: 'edit', mutations: [new UpdateTextMutation({ target: { slideId: 'ppt/slides/slide1.xml', origin: 'slide', elementId: '7' }, value: 'after' })] });
  const saving = session.save();
  expect(onListenerError).toHaveBeenCalledTimes(2);
  session.dispose(); session.dispose();
  resolve({ status: 'confirmed' });
  await expect(saving).resolves.toMatchObject({ status: 'confirmed' });
  expect(listener).toHaveBeenCalledTimes(2);
  expect(() => session.getSnapshot()).toThrow();
  expect(() => session.undo()).toThrow();
});

it('rejects invalid local commands atomically without changing history', () => {
  const { session, target, presentation, sendBatch } = setup();
  expect(() => session.apply({ id: 'bad', mutations: [
    new UpdateTextMutation({ target, value: 'not committed locally' }),
    new UpdateTextMutation({ target: { ...target, elementId: 'missing' }, value: 'bad' }),
  ] })).toThrow();
  expect(session.getSnapshot()).toMatchObject({ dirty: false, undoDepth: 0 });
  expect(session.getSnapshot().presentation).toBe(presentation);
  expect(sendBatch).not.toHaveBeenCalled();
});

it('retains the draft when preparation fails before sending', async () => {
  const { presentation, target } = setup();
  const sendBatch = vi.fn();
  const session = new PptxEditorSession({ presentation, sendBatch, createSaveId: () => { throw new Error('id unavailable'); } });
  session.apply({ id: 'edit', mutations: [new UpdateTextMutation({ target, value: 'draft' })] });
  await expect(session.save()).rejects.toThrow('id unavailable');
  expect(session.getSnapshot()).toMatchObject({ dirty: true, canEdit: true, saveStatus: 'failed' });
  expect(sendBatch).not.toHaveBeenCalled();
});

it('handles a long local history without transport fan-out', async () => {
  const { session, edit, sendBatch } = setup();
  for (let index = 0; index < 1000; index++) edit(String(index));
  for (let index = 0; index < 500; index++) session.undo();
  for (let index = 0; index < 250; index++) session.redo();
  expect(sendBatch).not.toHaveBeenCalled();
  expect(session.getSnapshot()).toMatchObject({ undoDepth: 750, redoDepth: 250 });
  await session.save();
  expect(sendBatch).toHaveBeenCalledTimes(1);
  expect(sendBatch.mock.calls[0][0].commands).toHaveLength(750);
  expect(textOf(session.getSnapshot().presentation)[0]).toBe('749');
});
