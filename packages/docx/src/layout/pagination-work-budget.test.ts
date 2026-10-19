import { describe, expect, it } from 'vitest';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { noteContinuationModel } from '../testing/note-continuation-model.js';
import { createLayoutBlockRepository, type LayoutStoryBlock, type LayoutParagraphBlock, type LayoutTableBlock } from './layout-source-store.js';
import type { SourceRef } from './types.js';
import { createPaginationWorkBudget, MAX_PAGINATION_ACQUISITION_MISSES } from './pagination-work-budget.js';

const ref = (storyInstance: string, story: SourceRef['story'] = 'textbox'): SourceRef => ({ story, storyInstance, path: [] });
const p = (...textBoxInputs: object[]): LayoutParagraphBlock => ({ type: 'paragraph',
  runs: textBoxInputs.map(textBoxInput => ({ type: 'shape', textBoxInput })),
} as unknown as LayoutParagraphBlock);
const box = (source: SourceRef) => ({ kind: 'complete', source });
const table = (...content: LayoutStoryBlock[]): LayoutTableBlock => ({ type: 'table', rows: [{ cells: [{ content }] }] } as unknown as LayoutTableBlock);
const header = ref('default', 'header');
const footer = ref('default', 'footer');

function fixture(stories: { source: SourceRef; body: LayoutStoryBlock[] }[],
  extra: Partial<Parameters<typeof createLayoutBlockRepository>[0]> = {}) {
  const base = layoutSourceStore(noteContinuationModel());
  const blocks = createLayoutBlockRepository({ body: Array(1000).fill(p()), stories,
    footnotes: [], endnotes: [], ...extra });
  return { ...base, blocks, bodyLayoutInput: { ...base.bodyLayoutInput,
    initialSection: { ...base.bodyLayoutInput.initialSection,
      headers: { default: header, first: null, even: null },
      footers: { default: stories.some(s => s.source.story === 'footer') ? footer : null, first: null, even: null },
    }, sequence: [],
  } };
}

describe('pagination source work attribution', () => {
  it('credits all recursive table, complete and compatibility text-box work once per selected page/owner/root', () => {
    const inner = ref('inner');
    const outer = ref('outer');
    const source = fixture([
      { source: header, body: [p(), table(p(box(outer)))] }, // 4 + outer(10)
      { source: footer, body: [p(box(outer))] }, // 2 + outer(10)
      { source: outer, body: [table(p(box(inner), { kind: 'compatibility', paragraphs: [{}, {}] }))] }, // 6 + inner(4)
      { source: inner, body: [p(), p(), p()] },
    ]);
    const budget = createPaginationWorkBudget(source);
    const initial = 32 * (1001 + 6 + 4);
    expect(budget.limit).toBe(initial);
    budget.admitPageStories('owner', 0, [header, header]);
    expect(budget.limit).toBe(initial + 32 * 14);
    budget.admitPageStories('owner', 0, [header, footer]);
    expect(budget.limit).toBe(initial + 32 * (14 + 12));
    budget.admitPageStories('owner', 0, [footer, header]);
    expect(budget.limit).toBe(initial + 32 * (14 + 12));
    budget.admitPageStories('owner', 1, [header]);
    budget.admitPageStories('other-owner', 0, [header]);
    expect(budget.limit).toBe(initial + 32 * (3 * 14 + 12));
  });

  it('counts body, footnote, endnote and separator definitions and their detached boxes once in U', () => {
    const bodyBox = ref('body-box');
    const noteBox = ref('note-box');
    const separator = ref('reserved:separator', 'footnote');
    const source = fixture([
      { source: header, body: [] },
      { source: bodyBox, body: [p(), p()] },
      { source: noteBox, body: [table(p())] },
    ], {
      body: [...Array(1000).fill(p()), p(box(bodyBox))],
      footnotes: [{ id: '1', content: [p(box(noteBox))] }],
      endnotes: [{ id: '2', content: [table(p(box(noteBox)))] }],
      reservedNoteStories: [{ source: separator, body: [p()] }],
    });
    // 1002 body + 2 footnote + 3 endnote + 2 separator + 3/3 detached
    // definitions. Referencing the same note box twice does not duplicate U.
    const budget = createPaginationWorkBudget(source);
    expect(budget.limit).toBe(32 * 1015);
    budget.admitPageStories('owner', 0, [bodyBox, noteBox]);
    expect(budget.limit).toBe(32 * 1015); // Detached boxes cannot mint W themselves.
    budget.admitPageStories('owner', 0, [ref('1', 'footnote'), ref('2', 'endnote'), separator]);
    expect(budget.limit).toBe(32 * (1015 + 5 + 6 + 2));
    budget.admitPageStories('owner', 0, [ref('1', 'footnote'), ref('2', 'endnote'), separator]);
    expect(budget.limit).toBe(32 * (1015 + 5 + 6 + 2));
  });

  it('rejects cyclic detached story references before acquisition', () => {
    const loop = ref('loop');
    const source = fixture([{ source: header, body: [p(box(loop))] },
      { source: loop, body: [p(box(loop))] }]);
    expect(() => createPaginationWorkBudget(source)).toThrow('Cyclic complete text-box story source');
  });

  it('memoizes deep shared story graphs and saturates recursive occurrence weights without expanding them', () => {
    const depth = 6000;
    const stories = Array.from({ length: depth }, (_, i) => ({ source: ref(`box-${i}`),
      body: [p(...(i + 1 < depth ? [box(ref(`box-${i + 1}`))] : []))] }));
    const source = fixture([{ source: header, body: [p(box(ref('box-0')))] }, ...stories]);
    const reads = new Map<string, number>();
    const budget = createPaginationWorkBudget({ ...source, blocks: { ...source.blocks,
      storyRoot(root) {
        reads.set(root.storyInstance, (reads.get(root.storyInstance) ?? 0) + 1);
        return source.blocks.storyRoot(root);
      },
    } });
    budget.admitPageStories('owner', 0, [header]);
    expect(budget.limit).toBe(32 * (1001 + 2 * depth + 2 + 2 * depth));
    expect(Math.max(...reads.values())).toBe(1);

    // Repeated references represent source occurrences, but definitions are
    // scanned once. Exponential logical work must saturate, never overflow.
    const branching = fixture([{ source: header, body: [p(box(ref('b0')))] },
      ...Array.from({ length: 24 }, (_, i) => ({ source: ref(`b${i}`),
        body: [p(...(i < 23 ? [box(ref(`b${i + 1}`)), box(ref(`b${i + 1}`))] : []))] }))]);
    const capped = createPaginationWorkBudget(branching);
    capped.admitPageStories('owner', 0, [header]);
    expect(capped.limit).toBe(MAX_PAGINATION_ACQUISITION_MISSES);
  });
});
