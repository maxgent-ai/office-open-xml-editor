import { afterAll, beforeAll, expect, it, vi } from 'vitest';
import type { Presentation, ShapeElement } from '@maxgent/ooxml/pptx';
import {
  PptxEditorSession, AddElementMutation, InsertSlideMutation, RemoveElementMutation,
  RemoveSlideMutation, UpdateShapeMutation, UpdateTextMutation, createElementRef,
  type ElementRef, type OfficeCliBatch,
} from '../../src';
import { resolveInheritedStylePatch } from '../../src/mutations/update-text/text-editing';
import {
  assertLiveOfficeCli, createLiveWorkspace, destroyLiveWorkspace, createDeck,
  addSlide, addShape, flushDeck, parseDeck, runBatch,
} from './harness';

let dir: string;
let path: string;
beforeAll(() => {
  assertLiveOfficeCli();
  dir = createLiveWorkspace('manual-save');
  path = `${dir}/deck.pptx`;
  createDeck(path);
  addSlide(path);
  for (const [id, text] of [['7', 'before'], ['8', 'other']]) {
    addShape(path, '/slide[1]', { id, text, x: '914400emu', y: '457200emu', width: '1828800emu', height: '914400emu' });
  }
  flushDeck(path);
});
afterAll(() => { if (dir) destroyLiveWorkspace(dir, [path]); });

// Compare editable content, geometry and order. Package part names, relationship
// ids and newly inserted slide ids are intentionally backend-owned metadata.
function editable(presentation: Presentation) {
  return presentation.slides.map((slide) => slide.elements.map((element) => {
    const shape = element as ShapeElement;
    return {
      id: shape.id, x: shape.x, y: shape.y, width: shape.width, height: shape.height,
      rotation: shape.rotation, flipH: shape.flipH, flipV: shape.flipV,
      text: shape.textBody?.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.type === 'text' ? run.text : '').join('')).join('\n'),
    };
  }));
}

it('replays local edits, dependencies, structural edits and saved undo in real batches', async () => {
  const presentation = parseDeck(path);
  let id = 0;
  const sendBatch = vi.fn(async (batch: OfficeCliBatch) => {
    runBatch(path, batch);
    return { status: 'confirmed' as const };
  });
  const session = new PptxEditorSession({ presentation, sendBatch, createSaveId: () => `save-${++id}` });
  const target = createElementRef(presentation.slides[0], presentation.slides[0].elements[0], 0);
  const edit = (target: ElementRef, value: string) => session.apply({
    id: `edit-${value}`, mutations: [new UpdateTextMutation({ target, value })],
  });
  edit(target, 'one'); edit(target, 'two');
  session.apply({ id: 'style', mutations: [new UpdateTextMutation({ target, style: { bold: true, fontSize: 24 } })] });
  session.undo(); session.redo();
  expect(sendBatch).not.toHaveBeenCalled();
  await session.save();
  expect(sendBatch).toHaveBeenCalledTimes(1);
  expect(editable(parseDeck(path))).toEqual(editable(session.getSnapshot().presentation));
  const savedShape = parseDeck(path).slides[0].elements[0] as ShapeElement;
  expect(savedShape.textBody?.paragraphs[0].runs[0]).toMatchObject({ bold: true, fontSize: 24 });

  // An undo after confirmation becomes the inverse in the next explicit save.
  session.undo(); await session.save();
  const undoneShape = parseDeck(path).slides[0].elements[0] as ShapeElement;
  const localUndone = session.getSnapshot().presentation.slides[0].elements[0] as ShapeElement;
  const localBody = localUndone.textBody as NonNullable<ShapeElement['textBody']>;
  const localParagraph = localBody.paragraphs[0];
  const localRun = localParagraph.runs[0];
  if (localRun.type !== 'text') throw new Error('Expected editable text');
  // Existing translator policy writes inherited values explicitly. Compare
  // effective style, not null-vs-explicit OOXML representation.
  expect(undoneShape.textBody?.paragraphs[0].runs[0]).toMatchObject(
    resolveInheritedStylePatch({ bold: localRun.bold, fontSize: localRun.fontSize }, {
      paragraph: localParagraph, textBody: localBody,
    }),
  );
  expect(editable(parseDeck(path))).toEqual(editable(session.getSnapshot().presentation));
  session.redo(); await session.save();

  const added = { ...target, elementId: '99' };
  session.apply({ id: 'add', mutations: [new AddElementMutation({
    target: added, element: { ...(presentation.slides[0].elements[0] as ShapeElement), id: '99' }, presentationElementIndex: 1,
  })] });
  edit(added, 'added');
  session.apply({ id: 'move', mutations: [new UpdateShapeMutation({ target: added, value: { x: 1828800, y: 914400 } })] });
  session.apply({ id: 'delete', mutations: [new RemoveElementMutation({ target })] });
  edit({ ...target, elementId: '8' }, 'survivor');
  await session.save();
  expect(editable(parseDeck(path))).toEqual(editable(session.getSnapshot().presentation));

  session.apply({ id: 'insert', mutations: [new InsertSlideMutation({ target: { slideId: 'local:new' }, index: 0 })] });
  edit(added, 'now second slide');
  const onNewSlide = { ...added, slideId: 'local:new' };
  session.apply({ id: 'add-on-new', mutations: [new AddElementMutation({
    target: onNewSlide, element: { ...(presentation.slides[0].elements[0] as ShapeElement), id: '99' }, presentationElementIndex: 0,
  })] });
  session.apply({ id: 'delete-slide', mutations: [new RemoveSlideMutation({ target })] });
  edit(onNewSlide, 'remaining slide');
  edit(onNewSlide, 'discarded'); session.undo();
  await session.save();
  expect(sendBatch).toHaveBeenCalledTimes(5);
  expect(editable(parseDeck(path))).toEqual(editable(session.getSnapshot().presentation));
  expect(session.getSnapshot().dirty).toBe(false);
});
