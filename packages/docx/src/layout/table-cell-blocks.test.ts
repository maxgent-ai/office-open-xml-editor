import { describe, expect, it } from 'vitest';
import type { CellElement, DocTable, DocTableCell } from '../types.js';
import type { FlowFragment } from './flow-fragment.js';
import type { ParagraphLayout, TableLayout } from './types.js';
import { acquireTableCellBlocks, resolveRetainedCellBlockPlacement } from './table-cell-blocks.js';

function paragraph(text: string): Extract<CellElement, { type: 'paragraph' }> {
  return { type: 'paragraph', runs: [{ type: 'text', text }] } as unknown as Extract<
    CellElement,
    { type: 'paragraph' }
  >;
}

function cell(content: DocTableCell['content']): DocTableCell {
  return { content, colSpan: 1, vMerge: null } as unknown as DocTableCell;
}

function table(cells: DocTableCell[]): DocTable {
  return {
    rows: [{ cells }], colWidths: [40], cellMarginLeft: 0, cellMarginRight: 0,
  } as unknown as DocTable;
}

function retainedParagraph(path: readonly number[], yPt: number): ParagraphLayout {
  return {
    kind: 'paragraph', id: path.join('.'),
    source: { story: 'body', storyInstance: 'body', path: [...path] },
    flowDomainId: 'table-cell', ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt, widthPt: 40, heightPt: 10 },
    inkBounds: { xPt: 0, yPt, widthPt: 10, heightPt: 8 },
    advancePt: 10, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

function retainedNestedTable(
  path: readonly number[],
  widthPt: number,
  heightPt: number,
  nestedBlocks: readonly FlowFragment[],
): TableLayout {
  const blocks = nestedBlocks.map((layout) => ({
    layout,
    offsetPt: 0,
    advancePt: layout.advancePt,
  }));
  return {
    kind: 'table', id: `nested-${path.join('.')}`,
    source: { story: 'body', storyInstance: 'body', path: [...path] },
    flowDomainId: 'table-cell', ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
    inkBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
    advancePt: heightPt, columnWidthsPt: [widthPt], borders: [],
    rows: [{
      kind: 'table-row', id: `nested-row-${path.join('.')}`,
      source: { story: 'body', storyInstance: 'body', path: [...path, 0] },
      flowDomainId: 'table-cell', ordinaryFlow: true,
      flowBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
      inkBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
      advancePt: heightPt, heightPt, contentHeightPt: heightPt,
      cells: [{
        kind: 'table-cell', id: `nested-cell-${path.join('.')}`,
        source: { story: 'body', storyInstance: 'body', path: [...path, 0, 0] },
        flowDomainId: 'table-cell', ordinaryFlow: true,
        flowBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
        inkBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
        contentBounds: { xPt: 0, yPt: 0, widthPt, heightPt },
        advancePt: heightPt, verticalMerge: 'none', vAlign: 'top', blocks,
      }],
    }],
  };
}

describe('table cell block acquisition', () => {
  it('retains paragraph/nested-table document order and recursively advances cell state', () => {
    const nested = table([cell([paragraph('nested')])]);
    const outerCell = cell([
      paragraph('before'),
      { type: 'table', ...nested } as DocTableCell['content'][number],
      paragraph('after'),
    ]);
    const outer = table([outerCell]);
    const events: string[] = [];

    const blocks = acquireTableCellBlocks({
      cell: outerCell, table: outer, cellTotalWidthPt: 50,
      outerState: { yPt: 99 }, sourcePath: [0, 0],
    }, {
      resolveContentWidthPt: (_cell, _table, totalWidthPt) => totalWidthPt - 10,
      createCellState: (_outerState, contentWidthPt) => ({ yPt: 0, contentWidthPt }),
      acquireParagraph: (state, _paragraph, contentWidthPt, path) => {
        events.push(`p:${path.join('.')}:${state.yPt}:${contentWidthPt}`);
        return retainedParagraph(path, state.yPt);
      },
      acquireNestedTable: (state, nestedTable, contentWidthPt, sourcePath, continuation, recurse) => {
        events.push(`t:${state.yPt}:${contentWidthPt}`);
        const nestedCell = nestedTable.rows[0]!.cells[0]!;
        expect(continuation).toEqual({ fromPrevious: false, onNext: false });
        return retainedNestedTable(sourcePath, contentWidthPt, 10, recurse(
          nestedCell,
          nestedTable,
          contentWidthPt,
          state,
          [...sourcePath, 0, 0],
        ));
      },
      advanceState: (state, advancePt) => { state.yPt += advancePt; },
    });

    expect(blocks.map((block) => block.kind)).toEqual(['paragraph', 'table', 'paragraph']);
    const nestedBlock = blocks[1];
    expect(nestedBlock?.kind === 'table' && 'flowBounds' in nestedBlock
      ? nestedBlock.rows[0]?.cells[0]?.blocks.map((block) => block.layout.kind)
      : []).toEqual(['paragraph']);
    expect(events).toEqual([
      'p:0.0.0:0:40',
      't:10:40',
      'p:0.0.1.0.0.0:0:30',
      'p:0.0.2:10:40',
    ]);
  });

  it('resolves equal paragraph-border adjacency in outer and nested cells before acquisition', () => {
    const borders = {
      top: { style: 'single', width: 1, space: 1, color: '111111' },
      bottom: { style: 'single', width: 1, space: 1, color: '111111' },
      left: { style: 'single', width: 1, space: 1, color: '111111' },
      right: { style: 'single', width: 1, space: 1, color: '111111' },
      between: { style: 'dashed', width: 1, space: 0, color: '222222' },
    } as const;
    const bordered = (value: string) => ({ ...paragraph(value), borders });
    const nested = table([cell([bordered('nested-a'), bordered('nested-b')])]);
    const outerCell = cell([
      bordered('outer-a'), bordered('outer-b'),
      { type: 'table', ...nested } as DocTableCell['content'][number],
    ]);
    const outer = table([outerCell]);
    const decisions: Array<readonly [string, unknown]> = [];

    acquireTableCellBlocks({
      cell: outerCell, table: outer, cellTotalWidthPt: 40,
      outerState: { yPt: 0 }, sourcePath: [0, 0],
    }, {
      resolveContentWidthPt: (_cell, _table, widthPt) => widthPt,
      createCellState: () => ({ yPt: 0 }),
      acquireParagraph: (state, para, _widthPt, path, edges?: unknown) => {
        decisions.push([(para.runs[0] as { text: string }).text, edges]);
        return retainedParagraph(path, state.yPt);
      },
      acquireNestedTable: (state, nestedTable, widthPt, sourcePath, continuation, recurse) => {
        expect(continuation).toEqual({ fromPrevious: false, onNext: false });
        return retainedNestedTable(sourcePath, widthPt, 20, recurse(
          nestedTable.rows[0]!.cells[0]!, nestedTable, widthPt, state,
          [...sourcePath, 0, 0],
        ));
      },
      advanceState: (state, advancePt) => { state.yPt += advancePt; },
    });

    expect(decisions).toEqual([
      ['outer-a', { top: 'top', bottom: 'none' }],
      ['outer-b', { top: 'between', bottom: 'bottom' }],
      ['nested-a', { top: 'top', bottom: 'none' }],
      ['nested-b', { top: 'between', bottom: 'bottom' }],
    ]);
  });
});

describe('split cell retained edge ownership', () => {
  it.each([
    ['top', 0],
    ['center', 40],
    ['bottom', 80],
  ] as const)('%s continuation does not repeat authored before/after spacing', (vAlign, expected) => {
    const source = {
      ...paragraph('continued'),
      spaceBefore: 12,
      spaceAfter: 8,
    } as Extract<CellElement, { type: 'paragraph' }>;
    const owner = {
      ...cell([source]),
      vAlign,
      marginTop: 0,
      marginBottom: 0,
    } as DocTableCell;
    const fragment = {
      ...retainedParagraph([0, 0, 0], 0),
      flowBounds: { xPt: 0, yPt: 0, widthPt: 40, heightPt: 20 },
      inkBounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: 20 },
      advancePt: 20,
      spacing: { beforePt: 0, afterPt: 0 },
      continuation: {
        lineStart: 1,
        lineEnd: 3,
        continuesFromPrevious: true,
        continuesOnNext: true,
      },
    } as ParagraphLayout;

    const placement = resolveRetainedCellBlockPlacement(owner, table([owner]), [fragment], 100);

    expect(placement.blockPlacements).toEqual([{ offsetPt: 0, advancePt: 20 }]);
    expect(placement.inkBlock).toEqual({ topPt: 0, heightPt: 20 });
    expect(placement.contentTranslationPt).toBe(expected);
  });

  it('final continuation owns only its retained trailing edge, not the source leading edge', () => {
    const source = { ...paragraph('final'), spaceBefore: 12, spaceAfter: 8 } as Extract<
      CellElement,
      { type: 'paragraph' }
    >;
    const owner = { ...cell([source]), vAlign: 'bottom' } as DocTableCell;
    const fragment = {
      ...retainedParagraph([0], 0),
      flowBounds: { xPt: 0, yPt: 0, widthPt: 40, heightPt: 28 },
      advancePt: 28,
      spacing: { beforePt: 0, afterPt: 8 },
      continuation: {
        lineStart: 2,
        lineEnd: 4,
        continuesFromPrevious: true,
        continuesOnNext: false,
      },
    } as ParagraphLayout;

    const placement = resolveRetainedCellBlockPlacement(owner, table([owner]), [fragment], 100);
    expect(placement.blockPlacements).toEqual([{ offsetPt: 0, advancePt: 20 }]);
    expect(placement.contentTranslationPt).toBe(80);
  });
});

describe('retained cell paragraph spacing', () => {
  it.each([
    ['ordinary max collapse', false, false, 12],
    ['previous contextual side', true, false, 2],
    ['current contextual side', false, true, 10],
    ['both contextual sides', true, true, 0],
  ] as const)('%s uses the shared paragraph-gap authority', (
    _name,
    previousContextual,
    currentContextual,
    expectedGap,
  ) => {
    const previous = Object.assign(paragraph('previous'), {
      styleId: 'shared',
      contextualSpacing: previousContextual,
    });
    const current = Object.assign(paragraph('current'), {
      styleId: 'shared',
      contextualSpacing: currentContextual,
    });
    const owner = cell([previous, current]);
    const previousLayout = {
      ...retainedParagraph([0], 0),
      advancePt: 30,
      spacing: { beforePt: 4, afterPt: 10 },
    };
    const currentLayout = {
      ...retainedParagraph([1], 0),
      advancePt: 32,
      spacing: { beforePt: 12, afterPt: 0 },
    };

    const placement = resolveRetainedCellBlockPlacement(
      owner,
      table([owner]),
      [previousLayout, currentLayout],
      100,
    );

    expect(placement.blockPlacements).toEqual([
      { offsetPt: 4, advancePt: 16 },
      { offsetPt: 20 + expectedGap, advancePt: 20 },
    ]);
  });

  it('resets paragraph spacing ownership across a nested table block', () => {
    const before = paragraph('before');
    const after = paragraph('after');
    const owner = cell([
      before,
      { type: 'table', ...table([cell([paragraph('nested')])]) } as CellElement,
      after,
    ]);
    const nested = retainedNestedTable([1], 40, 15, []);
    const placement = resolveRetainedCellBlockPlacement(owner, table([owner]), [
      { ...retainedParagraph([0], 0), advancePt: 20, spacing: { beforePt: 0, afterPt: 8 } },
      nested,
      { ...retainedParagraph([2], 0), advancePt: 22, spacing: { beforePt: 12, afterPt: 0 } },
    ], 100);

    expect(placement.blockPlacements).toEqual([
      { offsetPt: 0, advancePt: 12 },
      { offsetPt: 20, advancePt: 15 },
      { offsetPt: 47, advancePt: 10 },
    ]);
  });
});
