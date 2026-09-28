import { describe, expect, it } from 'vitest';
import { projectBodyOccurrence, translateBodyOccurrence } from './occurrence-projection.js';
import type {
  FloatingTablePlacementLayout, FloatingTablePositionInput, ParagraphLayout,
  ResolvedFloatingTablePlacementLayout, TableLayout, TextBoxLayout, TextPlacement,
} from './types.js';
import type { TableFragmentLayout, TableRowFragmentLayout } from './table-pagination.js';
import type { AnchorFrameResult } from './anchor-frame.js';
import { floatingTableAxesFollowHostFlow } from './retained-geometry-translation.js';
import { deepFreezePlainData, snapshotPlainData } from './plain-data.js';

const source = (path: readonly number[]) => ({ story: 'body' as const, storyInstance: 'body', path });
const rect = (xPt: number, yPt: number, widthPt = 10, heightPt = 10) => ({ xPt, yPt, widthPt, heightPt });

function paragraph(id = 'paragraph'): ParagraphLayout {
  return {
    kind: 'paragraph', id, source: source([0]), flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: rect(1, 2, 40, 12), inkBounds: rect(1, 2, 30, 10), advancePt: 12,
    spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [{
      range: { start: 0, end: 4 }, bounds: rect(1, 2, 20, 10), baselinePt: 10, advancePt: 12,
      placements: [{
        kind: 'text', text: 'PAGE', range: { start: 0, end: 4 }, origin: { xPt: 1, yPt: 10 },
        bounds: rect(1, 2, 20, 10), advancePt: 20, clusters: [], paintOps: [],
        color: { kind: 'explicit', color: '#000000' },
        fontRoute: { familyList: 'serif', scope: 'native', fingerprint: 'serif' },
        fontSizePt: 10, fontWeight: 400, fontStyle: 'normal', direction: 'ltr',
        decorations: [], dependency: 'page', sourceRunIndex: 7,
      } satisfies TextPlacement],
    }],
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

function table(id = 'table'): TableLayout {
  const child = paragraph(`${id}-cell-paragraph`);
  return {
    kind: 'table', id, source: source([1]), flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: rect(5, 6, 100, 20), inkBounds: rect(5, 6, 100, 20), advancePt: 20,
    columnWidthsPt: [100], borders: [], rows: [{
      kind: 'table-row', id: `${id}-row`, source: source([1, 0]), flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: rect(5, 6, 100, 20), inkBounds: rect(5, 6, 100, 20), advancePt: 20,
      heightPt: 20, contentHeightPt: 12, cells: [{
        kind: 'table-cell', id: `${id}-cell`, source: source([1, 0, 0]), flowDomainId: 'cell-source', ordinaryFlow: true,
        flowBounds: rect(5, 6, 100, 20), inkBounds: rect(5, 6, 100, 20), advancePt: 20,
        contentBounds: rect(7, 8, 96, 16), verticalMerge: 'none', vAlign: 'top',
        blocks: [{ layout: child, offsetPt: 2, advancePt: 12 }],
      }],
    }],
  };
}

const options = {
  occurrenceId: 'page 2/header',
  destination: {
    coordinateSpace: 'logical-page-points' as const,
    flowDomainId: 'body/page-2', translation: { xPt: 20, yPt: 30 },
  },
};

const floatingPosition = (overrides: Partial<FloatingTablePositionInput> = {}): FloatingTablePositionInput => ({
  leftFromTextPt: 1, rightFromTextPt: 1, topFromTextPt: 1, bottomFromTextPt: 1,
  horzAnchor: 'page', horzSpecified: true, vertAnchor: 'text', xPt: 0, yPt: 0,
  ...overrides,
});

type ProjectedTableGraph = TableLayout & Readonly<{
  floatingTables?: readonly FloatingTablePlacementLayout[];
  resolvedFloatingTables?: readonly ResolvedFloatingTablePlacementLayout[];
}>;

interface ProjectedGraphFacts {
  readonly nodeIds: ReadonlySet<string>;
  readonly anchorIds: ReadonlySet<string>;
  readonly occurrenceIds: ReadonlySet<string>;
}

function validateProjectedGraph(
  root: ParagraphLayout | TableLayout,
  outerOccurrenceId: string,
  externalAnchorRefs: ReadonlySet<string> = new Set(),
): ProjectedGraphFacts {
  const nodeOwners = new Map<string, Readonly<{ owner: object; kind: string }>>();
  const anchorOwners = new Map<string, object>();
  const nodeRefs: Array<Readonly<{ id: string; kind: 'drawing' | 'textbox' | 'table-cell' | 'table' }>> = [];
  const anchorRefs = new Set<string>();
  const floatOccurrenceIds = new Set<string>();
  const rowStamps = new Set<string>();
  const visited = new WeakSet<object>();
  const ownNode = (id: string, owner: object, kind: string) => {
    expect(id.startsWith(`${outerOccurrenceId}/node/`)).toBe(true);
    const prior = nodeOwners.get(id);
    expect(prior === undefined || prior.owner === owner).toBe(true);
    nodeOwners.set(id, { owner, kind });
  };
  const visit = (layout: ParagraphLayout | TableLayout): void => {
    if (visited.has(layout)) return;
    visited.add(layout);
    ownNode(layout.id, layout, layout.kind);
    if (layout.kind === 'paragraph') {
      for (const line of layout.lines) for (const placement of line.placements) {
        if (placement.kind === 'drawing') nodeRefs.push({ id: placement.drawingId, kind: 'drawing' });
        if (placement.kind === 'anchor-host' && placement.anchorOccurrenceId) {
          anchorRefs.add(placement.anchorOccurrenceId);
        }
      }
      for (const drawing of layout.drawings) {
        ownNode(drawing.id, drawing, 'drawing');
        drawing.textBoxIds?.forEach((id) => nodeRefs.push({ id, kind: 'textbox' }));
        if (drawing.anchorLayer) {
          const id = drawing.anchorLayer.occurrenceId;
          expect(id.startsWith(`${outerOccurrenceId}/anchor/`)).toBe(true);
          const prior = anchorOwners.get(id);
          expect(prior === undefined || prior === drawing).toBe(true);
          anchorOwners.set(id, drawing);
        }
      }
      for (const textBox of layout.textBoxes) {
        ownNode(textBox.id, textBox, 'textbox');
        textBox.story.blocks.forEach((block) => {
          if (block.kind === 'paragraph' || block.kind === 'table') visit(block);
        });
      }
      for (const exclusion of layout.exclusions) {
        ownNode(exclusion.id, exclusion, 'exclusion');
        if (exclusion.anchorOccurrenceId) anchorRefs.add(exclusion.anchorOccurrenceId);
      }
      layout.anchorFrames?.forEach((frame) => anchorRefs.add(frame.occurrenceId));
      return;
    }
    const tableGraph = layout as ProjectedTableGraph;
    for (const row of layout.rows) {
      ownNode(row.id, row, 'table-row');
      if ('occurrenceId' in row && typeof row.occurrenceId === 'string') {
        expect(row.occurrenceId.startsWith(`${outerOccurrenceId}/occurrence/`)).toBe(true);
        rowStamps.add(row.occurrenceId);
      }
      for (const cell of row.cells) {
        ownNode(cell.id, cell, 'table-cell');
        cell.blocks.forEach((block) => visit(block.layout));
      }
    }
    const validateSource = (placement: FloatingTablePlacementLayout) => {
      expect(placement.occurrenceId.startsWith(`${outerOccurrenceId}/occurrence/`)).toBe(true);
      floatOccurrenceIds.add(placement.occurrenceId);
      nodeRefs.push({ id: placement.hostCellId, kind: 'table-cell' });
      nodeRefs.push({ id: placement.tableId, kind: 'table' });
      expect(placement.tableId).toBe(placement.child.id);
      visit(placement.child);
    };
    tableGraph.floatingTables?.forEach(validateSource);
    tableGraph.resolvedFloatingTables?.forEach((resolved) => {
      validateSource(resolved.source);
      floatOccurrenceIds.add(resolved.occurrenceId);
      expect(resolved.occurrenceId).toBe(resolved.source.occurrenceId);
      expect(resolved.child).toBe(resolved.source.child);
    });
  };
  visit(root);
  for (const reference of nodeRefs) expect(nodeOwners.get(reference.id)?.kind).toBe(reference.kind);
  for (const reference of anchorRefs) {
    expect(anchorOwners.has(reference) || externalAnchorRefs.has(reference)).toBe(true);
  }
  return {
    nodeIds: new Set(nodeOwners.keys()),
    anchorIds: new Set(anchorOwners.keys()),
    occurrenceIds: new Set([...floatOccurrenceIds, ...rowStamps]),
  };
}

function expectDisjoint(left: ReadonlySet<string>, right: ReadonlySet<string>): void {
  for (const value of left) expect(right.has(value)).toBe(false);
}

it('uses one tblpPr predicate for mixed-axis host ownership', () => {
  expect(floatingTableAxesFollowHostFlow(floatingPosition({
    horzAnchor: 'page', horzSpecified: true, vertAnchor: 'text',
  }))).toEqual({ x: false, y: true });
  expect(floatingTableAxesFollowHostFlow(floatingPosition({
    horzAnchor: 'text', horzSpecified: true, vertAnchor: 'page',
  }))).toEqual({ x: true, y: false });
});

describe('translateBodyOccurrence', () => {
  it('translates geometry while preserving the complete identity graph', () => {
    const retained = paragraph();
    const translated = translateBodyOccurrence(retained, { xPt: 20, yPt: 30 });

    expect(translated).toMatchObject({
      id: retained.id, flowDomainId: retained.flowDomainId,
      flowBounds: { xPt: 21, yPt: 32 },
      lines: [{ baselinePt: 40, placements: [{ text: 'PAGE', origin: { xPt: 21, yPt: 40 } }] }],
    });
    expect(retained.flowBounds).toEqual(rect(1, 2, 40, 12));
  });

  it('rejects non-finite translations', () => {
    expect(() => translateBodyOccurrence(paragraph(), { xPt: Number.NaN, yPt: 0 })).toThrow(RangeError);
  });
});

describe('projectBodyOccurrence', () => {
  it('shares verified unchanged source data while isolating occurrence-owned geometry and identity', () => {
    const retained = snapshotPlainData(table(), 'table acquisition') as TableLayout;
    const first = projectBodyOccurrence(retained, options);
    const second = projectBodyOccurrence(retained, { ...options, occurrenceId: 'page-3' });

    expect(first.source).toBe(retained.source);
    expect(second.source).toBe(retained.source);
    expect(first.source.path).toBe(retained.source.path);
    expect(first.rows[0]).not.toBe(second.rows[0]);
    expect(first.rows[0]!.cells[0]!.blocks[0]!.layout)
      .not.toBe(second.rows[0]!.cells[0]!.blocks[0]!.layout);
    expect(first.id).not.toBe(second.id);
    expect(first.flowBounds).not.toBe(second.flowBounds);
    expect(Object.isFrozen(first.source)).toBe(true);
  });

  it('shares a deep-frozen paragraph acquisition payload while isolating each occurrence', () => {
    // Paragraph acquisition deep-freezes its result, and the acquisition memo
    // keeps it for later convergence passes. Pages must reference its
    // unchanged payload instead of holding a second copy of it.
    const base = paragraph();
    const text = base.lines[0]!.placements[0] as TextPlacement;
    const retained = deepFreezePlainData({
      ...base,
      lines: [{ ...base.lines[0]!, placements: [{
        ...text,
        clusters: [{ range: { start: 0, end: 4 }, offset: { xPt: 0, yPt: 0 }, advancePt: 20 }],
        paintOps: [{ text: 'PAGE', origin: { xPt: 0, yPt: 0 } }],
      } as unknown as TextPlacement] }],
    }) as ParagraphLayout;
    const first = projectBodyOccurrence(retained, options);
    const second = projectBodyOccurrence(retained, { ...options, occurrenceId: 'page-3' });
    const retainedText = retained.lines[0]!.placements[0] as TextPlacement;
    const firstText = first.lines[0]!.placements[0] as TextPlacement;
    const secondText = second.lines[0]!.placements[0] as TextPlacement;

    expect(firstText.clusters).toBe(retainedText.clusters);
    expect(firstText.paintOps).toBe(retainedText.paintOps);
    expect(firstText.color).toBe(retainedText.color);
    expect(secondText.paintOps).toBe(retainedText.paintOps);
    expect(first.spacing).toBe(retained.spacing);
    // Occurrence-owned geometry and identity stay distinct per placement.
    expect(firstText).not.toBe(secondText);
    expect(firstText.origin).toEqual({ xPt: 21, yPt: 40 });
    expect(retainedText.origin).toEqual({ xPt: 1, yPt: 10 });
    expect(first.id).not.toBe(second.id);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(firstText)).toBe(true);
    expect(Object.isFrozen(firstText.origin)).toBe(true);
    validateProjectedGraph(first, options.occurrenceId);
  });

  it('copies unchanged data from an unverified mutable source', () => {
    const retained = table();
    const projected = projectBodyOccurrence(retained, options);
    (retained.source.path as number[])[0] = 99;

    expect(projected.source.path).toEqual([1]);
    expect(projected.source).not.toBe(retained.source);
  });

  it('rejects two distinct drawing owners claiming one raw anchor occurrence ID', () => {
    const retained = paragraph();
    const drawing = (id: string) => ({
      kind: 'drawing' as const, id, source: source([0, id.length]), flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: rect(1, 2), inkBounds: rect(1, 2), advancePt: 0, commands: [],
      anchorLayer: { occurrenceId: 'same-anchor', behindDoc: false, relativeHeight: 0, sourceOrder: 0,
        horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const },
    });

    expect(() => projectBodyOccurrence({ ...retained, drawings: [drawing('one'), drawing('two')] }, options))
      .toThrow('duplicate anchor occurrence owner');
  });

  it('allows repeated references to one drawing anchor owner', () => {
    const retained = paragraph();
    const drawing = {
      kind: 'drawing' as const, id: 'shared-drawing', source: source([0, 1]), flowDomainId: 'body',
      ordinaryFlow: false, flowBounds: rect(1, 2), inkBounds: rect(1, 2), advancePt: 0, commands: [],
      anchorLayer: { occurrenceId: 'shared-anchor', behindDoc: false, relativeHeight: 0, sourceOrder: 0,
        horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const },
    };

    const projected = projectBodyOccurrence({ ...retained, drawings: [drawing, drawing] }, options);

    expect(projected.drawings[0]).toBe(projected.drawings[1]);
    validateProjectedGraph(projected, options.occurrenceId);
  });

  it('rejects two distinct floating source owners claiming one raw occurrence ID', () => {
    const firstChild = table('float-child-one');
    const secondChild = table('float-child-two');
    const placement = (child: TableLayout): FloatingTablePlacementLayout => ({
      kind: 'floating-table-placement', occurrenceId: 'same-float', ownership: 'source',
      physicalPageIndex: 0, displayPageNumber: 1, hostCellId: 'table-cell', sourceBlockIndex: 0,
      anchorBlockIndex: 0, tableId: child.id, overlap: 'overlap', positioning: floatingPosition(),
      anchorBounds: rect(1, 2), child,
    });
    const retained = Object.assign(table(), {
      floatingTables: [placement(firstChild), placement(secondChild)],
      resolvedFloatingTables: [] as ResolvedFloatingTablePlacementLayout[],
    });

    expect(() => projectBodyOccurrence(retained, options))
      .toThrow('duplicate floating placement occurrence owner');
  });

  it('re-keys a paragraph graph deterministically and preserves acquisition occurrence identity', () => {
    const retained = paragraph();
    const drawing = {
      kind: 'drawing' as const, id: 'drawing/a', source: source([0, 0]), flowDomainId: 'body', ordinaryFlow: false,
      flowBounds: rect(3, 4), inkBounds: rect(3, 4), advancePt: 0, commands: [],
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE' as const,
        severity: 'error' as const,
        source: source([0, 0]),
        message: 'Unsupported optional drawing omitted',
      }],
      anchorLayer: {
        occurrenceId: 'anchor/a', behindDoc: false, relativeHeight: 1, sourceOrder: 0,
        horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const,
      },
    };
    const withDrawing = {
      ...retained,
      drawings: [drawing],
      lines: [{ ...retained.lines[0]!, placements: [{
        kind: 'drawing' as const, range: { start: 0, end: 0 }, drawingId: drawing.id,
        bounds: drawing.flowBounds, advancePt: 0,
      }] }],
      exclusions: [{
        id: 'exclusion', wrap: 'square' as const, bounds: drawing.flowBounds,
        polygon: [{ xPt: 3, yPt: 4 }], anchorOccurrenceId: 'anchor/a',
      }],
    };

    const first = projectBodyOccurrence(withDrawing, options);
    const again = projectBodyOccurrence(withDrawing, options);
    const anchor = first.drawings[0]!.anchorLayer!;

    expect(first).toEqual(again);
    expect(first.id).toBe('page 2/header/node/paragraph');
    expect(first.drawings[0]!.id).toBe('page 2/header/node/drawing%2Fa');
    expect(first.lines[0]!.placements[0]).toMatchObject({ drawingId: first.drawings[0]!.id });
    expect(first.drawings[0]!.diagnostics).toEqual(drawing.diagnostics);
    expect(Object.isFrozen(first.drawings[0]!.diagnostics)).toBe(true);
    expect(Object.isFrozen(first.drawings[0]!.diagnostics![0])).toBe(true);
    expect(anchor).toMatchObject({
      occurrenceId: 'page 2/header/anchor/anchor%2Fa', acquisitionOccurrenceId: 'anchor/a',
    });
    expect(first.exclusions[0]!.anchorOccurrenceId).toBe(anchor.occurrenceId);
    validateProjectedGraph(first, options.occurrenceId);
  });

  it('makes repeated occurrences disjoint while preserving source and shaped field metadata', () => {
    const retained = paragraph();
    const first = projectBodyOccurrence(retained, options);
    const second = projectBodyOccurrence(retained, { ...options, occurrenceId: 'page-3' });

    expect(first.id).not.toBe(second.id);
    expect(first.source).toEqual(retained.source);
    expect(first.lines[0]!.placements[0]).toEqual(expect.objectContaining({ text: 'PAGE' }));
    expect(first.lines[0]!.placements[0]).toEqual(expect.objectContaining({
      text: 'PAGE', dependency: 'page', sourceRunIndex: 7,
    }));
    const firstFacts = validateProjectedGraph(first, options.occurrenceId);
    const secondFacts = validateProjectedGraph(second, 'page-3');
    expectDisjoint(firstFacts.nodeIds, secondFacts.nodeIds);
    expectDisjoint(firstFacts.anchorIds, secondFacts.anchorIds);
    expectDisjoint(firstFacts.occurrenceIds, secondFacts.occurrenceIds);
  });

  it('projects table, row and cell identity without double-translating cell-local blocks', () => {
    const retained = table();
    const projected = projectBodyOccurrence(retained, options);
    const cell = projected.rows[0]!.cells[0]!;

    expect(projected.flowBounds).toMatchObject({ xPt: 25, yPt: 36 });
    expect(projected.rows[0]!.id).toBe('page 2/header/node/table-row');
    expect(cell.id).toBe('page 2/header/node/table-cell');
    expect(cell.flowDomainId).toBe('body/page-2/occurrence/page%202%2Fheader/cell/table-cell');
    expect(cell.blocks[0]!.layout.flowBounds).toEqual(retained.rows[0]!.cells[0]!.blocks[0]!.layout.flowBounds);
    expect(cell.blocks[0]!.layout.id).toBe('page 2/header/node/table-cell-paragraph');
  });

  it('preserves resolved floating table aliases through snapshotting and keeps page-owned axes fixed', () => {
    const nested = table('nested');
    const sourcePlacement: FloatingTablePlacementLayout = {
      kind: 'floating-table-placement' as const, occurrenceId: 'float-source', ownership: 'source' as const,
      physicalPageIndex: 4, displayPageNumber: 9, hostCellId: 'table-cell', sourceBlockIndex: 2,
      anchorBlockIndex: 2, tableId: 'nested', overlap: 'overlap' as const,
      positioning: floatingPosition({ horzAnchor: 'page', horzSpecified: true, vertAnchor: 'margin' }),
      columnBounds: rect(60, 70), anchorBounds: rect(70, 80), child: nested,
    };
    const resolved: ResolvedFloatingTablePlacementLayout = {
      kind: 'resolved-floating-table-placement' as const, occurrenceId: 'float-source',
      xPt: 70, yPt: 80, bounds: rect(70, 80), exclusionBounds: rect(69, 79, 12, 12),
      overlap: 'overlap' as const, child: nested, source: sourcePlacement,
    };
    const retained = Object.assign(table(), {
      floatingTables: [] as FloatingTablePlacementLayout[], resolvedFloatingTables: [resolved],
    });

    const projected = projectBodyOccurrence(retained, options) as typeof retained;
    const projectedResolved = projected.resolvedFloatingTables[0]!;

    expect(projectedResolved.bounds).toEqual(rect(70, 80));
    expect(projectedResolved.source.anchorBounds).toEqual(rect(90, 110));
    expect(projectedResolved.source.columnBounds).toEqual(rect(80, 100));
    expect(projectedResolved.source.child.flowDomainId)
      .toBe('body/page-2/occurrence/page%202%2Fheader/cell/table-cell');
    expect(projectedResolved.child).toBe(projectedResolved.source.child);
    const clone = structuredClone(projectedResolved);
    expect(clone.child).toBe(clone.source.child);
    expect(projectedResolved.occurrenceId).toBe('page 2/header/occurrence/float-source');
    expect(projectedResolved.source.occurrenceId).toBe('page 2/header/occurrence/float-source');
    validateProjectedGraph(projected, options.occurrenceId);
  });

  it('translates an unresolved-only float anchor and column with host flow', () => {
    const child = table('unresolved-child');
    const unresolved: FloatingTablePlacementLayout = {
      kind: 'floating-table-placement', occurrenceId: 'unresolved', ownership: 'source',
      physicalPageIndex: 0, displayPageNumber: 1, hostCellId: 'table-cell', sourceBlockIndex: 0,
      anchorBlockIndex: 0, tableId: child.id, overlap: 'overlap',
      positioning: floatingPosition({ horzAnchor: 'page', horzSpecified: true, vertAnchor: 'page' }),
      anchorBounds: rect(5, 6), columnBounds: rect(7, 8), child,
    };
    const retained = Object.assign(table(), {
      floatingTables: [unresolved], resolvedFloatingTables: [] as ResolvedFloatingTablePlacementLayout[],
    });

    const translated = translateBodyOccurrence(retained, { xPt: 20, yPt: 30 });

    expect(translated.floatingTables[0]).toMatchObject({
      anchorBounds: { xPt: 25, yPt: 36 }, columnBounds: { xPt: 27, yPt: 38 },
      child: { flowBounds: { xPt: 25, yPt: 36 } },
    });
    validateProjectedGraph(projectBodyOccurrence(retained, options), options.occurrenceId);
  });

  it.each([
    {
      name: 'x-page/y-host', positioning: { horzAnchor: 'page', horzSpecified: true, vertAnchor: 'text' },
      expected: { xPt: 10, yPt: 50 }, child: { xPt: 5, yPt: 36 }, exclusion: { xPt: 9, yPt: 49 },
    },
    {
      name: 'x-host/y-page', positioning: { horzAnchor: 'text', horzSpecified: true, vertAnchor: 'page' },
      expected: { xPt: 30, yPt: 20 }, child: { xPt: 25, yPt: 6 }, exclusion: { xPt: 29, yPt: 19 },
    },
  ])('translates successful mixed-axis resolved float geometry: $name', ({ positioning, expected, child, exclusion }) => {
    const nested = table(`mixed-${positioning.horzAnchor}-${positioning.vertAnchor}`);
    const sourcePlacement: FloatingTablePlacementLayout = {
      kind: 'floating-table-placement', occurrenceId: `mixed-${positioning.horzAnchor}`, ownership: 'source',
      physicalPageIndex: 0, displayPageNumber: 1, hostCellId: 'table-cell', sourceBlockIndex: 0,
      anchorBlockIndex: 0, tableId: nested.id, overlap: 'overlap',
      positioning: floatingPosition(positioning), anchorBounds: rect(3, 4), child: nested,
    };
    const resolved: ResolvedFloatingTablePlacementLayout = {
      kind: 'resolved-floating-table-placement', occurrenceId: sourcePlacement.occurrenceId,
      xPt: 10, yPt: 20, bounds: rect(10, 20), exclusionBounds: rect(9, 19, 12, 12),
      overlap: 'overlap', child: nested, source: sourcePlacement,
    };
    const retained = Object.assign(table(), {
      floatingTables: [] as FloatingTablePlacementLayout[], resolvedFloatingTables: [resolved],
    });

    const translated = translateBodyOccurrence(retained, { xPt: 20, yPt: 30 });
    const result = translated.resolvedFloatingTables[0]!;

    expect(result).toMatchObject({ xPt: expected.xPt, yPt: expected.yPt,
      bounds: expected, exclusionBounds: exclusion, child: { flowBounds: child } });
    expect(result.child).toBe(result.source.child);
  });

  it('projects a final-frame resolved float to the page exactly once', () => {
    const nested = table('final-frame-child');
    const sourcePlacement: FloatingTablePlacementLayout = {
      kind: 'floating-table-placement', occurrenceId: 'final-frame', ownership: 'source',
      physicalPageIndex: 0, displayPageNumber: 1, hostCellId: 'table-cell', sourceBlockIndex: 0,
      anchorBlockIndex: 0, tableId: nested.id, overlap: 'overlap',
      positioning: floatingPosition({ horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text' }),
      anchorBounds: rect(23, 34), columnBounds: rect(20, 30), child: nested,
    };
    const resolved: ResolvedFloatingTablePlacementLayout = {
      kind: 'resolved-floating-table-placement', occurrenceId: sourcePlacement.occurrenceId,
      xPt: 25, yPt: 36, bounds: rect(25, 36), exclusionBounds: rect(24, 35, 12, 12),
      overlap: 'overlap', child: nested, source: sourcePlacement,
    };
    const retained = Object.assign(table(), {
      floatingTables: [] as FloatingTablePlacementLayout[],
      resolvedFloatingTables: [resolved],
      resolvedFloatingTableCoordinateSpace: 'logical-page-points' as const,
    });

    const projected = translateBodyOccurrence(retained, { xPt: 20, yPt: 30 });
    const result = projected.resolvedFloatingTables[0]!;

    expect(projected.flowBounds).toMatchObject({ xPt: 25, yPt: 36 });
    expect(result.bounds).toEqual(rect(25, 36));
    expect(result.exclusionBounds).toEqual(rect(24, 35, 12, 12));
    expect(result.source.anchorBounds).toEqual(rect(23, 34));
    expect(result.source.columnBounds).toEqual(rect(20, 30));
  });

  it('preserves PAGE and NUMPAGES dependency/source-run metadata verbatim', () => {
    const retained = paragraph();
    const page = retained.lines[0]!.placements[0] as TextPlacement;
    const totalPages: TextPlacement = {
      ...page, text: 'NUMPAGES', dependency: 'total-pages', sourceRunIndex: 8,
      range: { start: 4, end: 12 },
    };
    const graph = { ...retained, lines: [{ ...retained.lines[0]!, placements: [page, totalPages] }] };

    const projected = projectBodyOccurrence(graph, options);

    expect(projected.lines[0]!.placements).toEqual([
      expect.objectContaining({ text: 'PAGE', dependency: 'page', sourceRunIndex: 7 }),
      expect.objectContaining({ text: 'NUMPAGES', dependency: 'total-pages', sourceRunIndex: 8 }),
    ]);
  });

  it('returns a structured-clone-safe deep-frozen snapshot without changing input', () => {
    const retained = table();
    const before = structuredClone(retained);
    const projected = projectBodyOccurrence(retained, options);

    expect(structuredClone(projected)).toEqual(projected);
    expect(Object.isFrozen(projected)).toBe(true);
    expect(Object.isFrozen(projected.rows[0]!.cells[0]!.blocks[0]!.layout)).toBe(true);
    expect(retained).toEqual(before);
  });

  it('rejects empty identities and conflicting ownership of one table object', () => {
    expect(() => projectBodyOccurrence(paragraph(), { ...options, occurrenceId: '' })).toThrow(RangeError);
    expect(() => projectBodyOccurrence(paragraph(), {
      ...options, destination: { ...options.destination, flowDomainId: '' },
    })).toThrow(RangeError);
    const shared = table('shared');
    const retained = table();
    const conflicting = {
      ...retained,
      rows: [{ ...retained.rows[0]!, cells: [
        { ...retained.rows[0]!.cells[0]!, id: 'cell-a', blocks: [{ layout: shared, offsetPt: 0, advancePt: 20 }] },
        { ...retained.rows[0]!.cells[0]!, id: 'cell-b', blocks: [{ layout: shared, offsetPt: 0, advancePt: 20 }] },
      ] }],
    };

    expect(() => projectBodyOccurrence(conflicting, options)).toThrow('incompatible projection ownership');
  });

  it('allows fragment descendants to share their semantic page occurrence identity', () => {
    const retained = table();
    const drawing = {
      kind: 'drawing' as const, id: 'anchor-drawing', source: source([1, 0, 0, 0]),
      flowDomainId: 'cell-source', ordinaryFlow: false, flowBounds: rect(0, 0),
      inkBounds: rect(0, 0), advancePt: 0, commands: [],
      anchorLayer: {
        occurrenceId: 'duplicate-owner', behindDoc: false, relativeHeight: 0, sourceOrder: 0,
        horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const,
      },
    };
    const child = { ...paragraph('owned-child'), drawings: [drawing] };
    const conflicting = {
      ...retained,
      rows: [{
        ...retained.rows[0]!, occurrenceId: 'duplicate-owner',
        cells: [{ ...retained.rows[0]!.cells[0]!, blocks: [{ layout: child, offsetPt: 0, advancePt: 12 }] }],
      }],
    };

    const first = projectBodyOccurrence(conflicting, options);
    const secondOccurrence = { ...options, occurrenceId: 'page-3/cross-category' };
    const second = projectBodyOccurrence(conflicting, secondOccurrence);
    const firstFacts = validateProjectedGraph(first, options.occurrenceId);
    const secondFacts = validateProjectedGraph(second, secondOccurrence.occurrenceId);
    expectDisjoint(firstFacts.nodeIds, secondFacts.nodeIds);
    expectDisjoint(firstFacts.anchorIds, secondFacts.anchorIds);
    expectDisjoint(firstFacts.occurrenceIds, secondFacts.occurrenceIds);
  });

  it('preserves repeated-header and split fragment metadata while re-keying shared row occurrences', () => {
    const base = table();
    const fragmentRow = (
      id: string,
      logicalRowIndex: number,
      fragmentIndex: number,
      ownership: 'source' | 'repeated-header',
    ): TableRowFragmentLayout => ({
      ...base.rows[0]!, id, logicalRowIndex, fragmentIndex, ownership,
      occurrenceId: 'page-9', physicalPageIndex: 4, displayPageNumber: 9,
      repeatedHeader: ownership === 'repeated-header',
      cells: base.rows[0]!.cells.map((cell) => ({
        ...cell, id: `${id}-cell`, source: source([1, logicalRowIndex, 0]),
        blocks: cell.blocks.map((block) => ({
          ...block,
          layout: { ...block.layout, id: `${id}-paragraph`, source: source([1, logicalRowIndex, 0, 0]) },
        })),
        contentRanges: [{ kind: 'paragraph', blockIndex: 0, lineStart: 1, lineEnd: 2 }],
        ...(fragmentIndex > 0 ? { visualMergeOwnership: 'continuation' as const } : {}),
      })),
    });
    const fragment: TableFragmentLayout = {
      ...base,
      rows: [fragmentRow('header-row', 0, 0, 'repeated-header'), fragmentRow('split-row', 3, 2, 'source')],
      floatingTables: [], resolvedFloatingTables: [], resolvedFloatingTableCoordinateSpace: 'logical-page-points',
    };

    const projected = projectBodyOccurrence(fragment, options) as TableFragmentLayout;

    expect(projected.rows.map((row) => row.occurrenceId)).toEqual([
      'page 2/header/occurrence/page-9', 'page 2/header/occurrence/page-9',
    ]);
    expect(projected.rows.map((row) => ({
      logicalRowIndex: row.logicalRowIndex, fragmentIndex: row.fragmentIndex,
      ownership: row.ownership, physicalPageIndex: row.physicalPageIndex,
      displayPageNumber: row.displayPageNumber, ranges: row.cells[0]!.contentRanges,
      merge: row.cells[0]!.visualMergeOwnership,
    }))).toEqual([
      { logicalRowIndex: 0, fragmentIndex: 0, ownership: 'repeated-header', physicalPageIndex: 4,
        displayPageNumber: 9, ranges: [{ kind: 'paragraph', blockIndex: 0, lineStart: 1, lineEnd: 2 }], merge: undefined },
      { logicalRowIndex: 3, fragmentIndex: 2, ownership: 'source', physicalPageIndex: 4,
        displayPageNumber: 9, ranges: [{ kind: 'paragraph', blockIndex: 0, lineStart: 1, lineEnd: 2 }], merge: 'continuation' },
    ]);
    validateProjectedGraph(projected, options.occurrenceId);
  });

  it('preserves shared textbox paragraph aliases and rejects textbox cycles before recursion', () => {
    const child = paragraph('textbox-child');
    const textBox = (id: string, paragraphs: readonly ParagraphLayout[]): TextBoxLayout => ({
      kind: 'textbox', id, source: source([0, 1]), flowDomainId: 'textbox', ordinaryFlow: false,
      flowBounds: rect(5, 6), inkBounds: rect(5, 6), advancePt: 0,
      story: {
        story: 'textbox',
        flowBounds: rect(5, 6),
        inkBounds: rect(5, 6),
        blocks: paragraphs,
        advancePt: paragraphs.reduce((sum, paragraph) => sum + paragraph.advancePt, 0),
        diagnostics: [],
      },
      transform: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      writingMode: 'horizontal-tb', insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    });
    const aliased = { ...paragraph(), textBoxes: [textBox('box-a', [child, child])] };

    const projected = projectBodyOccurrence(aliased, options);

    expect(projected.textBoxes[0]!.story.blocks[0])
      .toBe(projected.textBoxes[0]!.story.blocks[1]);
    const cyclicChildren: ParagraphLayout[] = [];
    const cyclic = { ...paragraph('cyclic'), textBoxes: [textBox('cycle-box', cyclicChildren)] };
    cyclicChildren.push(cyclic);
    expect(() => projectBodyOccurrence(cyclic, options)).toThrow(/acyclic/);
  });

  it('rejects one table translated under incompatible mixed-axis ownership', () => {
    const shared = table('shared-float-child');
    const hostSource: FloatingTablePlacementLayout = {
      kind: 'floating-table-placement', occurrenceId: 'host-float', ownership: 'source',
      physicalPageIndex: 0, displayPageNumber: 1, hostCellId: 'cell', sourceBlockIndex: 0,
      anchorBlockIndex: 0, tableId: shared.id, overlap: 'overlap',
      positioning: floatingPosition({ horzAnchor: 'text', horzSpecified: true, vertAnchor: 'text' }),
      anchorBounds: rect(1, 2), child: shared,
    };
    const pageSource: FloatingTablePlacementLayout = {
      ...hostSource, occurrenceId: 'page-float', positioning: floatingPosition({
        horzAnchor: 'page', horzSpecified: true, vertAnchor: 'text',
      }),
    };
    const resolved: ResolvedFloatingTablePlacementLayout = {
      kind: 'resolved-floating-table-placement', occurrenceId: pageSource.occurrenceId,
      xPt: 1, yPt: 2, bounds: rect(1, 2), exclusionBounds: rect(0, 1, 12, 12),
      overlap: 'overlap', child: shared, source: pageSource,
    };
    const retained = Object.assign(table(), {
      floatingTables: [hostSource, pageSource], resolvedFloatingTables: [resolved],
    });

    expect(() => translateBodyOccurrence(retained, { xPt: 20, yPt: 30 }))
      .toThrow('incompatible projection ownership');
  });

  it('validates every retained ID reference against its live occurrence-local owner', () => {
    const retained = paragraph();
    const drawing = {
      kind: 'drawing' as const, id: 'drawing-live', source: source([0, 2]), flowDomainId: 'body',
      ordinaryFlow: false, flowBounds: rect(1, 2), inkBounds: rect(1, 2), advancePt: 0, commands: [],
      anchorLayer: { occurrenceId: 'anchor-live', behindDoc: false, relativeHeight: 0, sourceOrder: 0,
        horizontalOwnership: 'host' as const, verticalOwnership: 'host' as const },
    };
    const graph = {
      ...retained, drawings: [drawing],
      lines: [{ ...retained.lines[0]!, placements: [{ kind: 'drawing' as const,
        range: { start: 0, end: 0 }, drawingId: drawing.id, bounds: drawing.flowBounds, advancePt: 0 }] }],
      exclusions: [{ id: 'wrap-live', wrap: 'square' as const, bounds: rect(1, 2), polygon: [],
        anchorOccurrenceId: drawing.anchorLayer.occurrenceId }],
    };
    const projected = projectBodyOccurrence(graph, options);
    validateProjectedGraph(projected, options.occurrenceId);
  });

  it('translates anchor frames and line numbers while keeping vertical textbox children local', () => {
    const axis = (axisName: 'horizontal' | 'vertical', origin: number) => ({
      axis: axisName, status: 'resolved' as const, relativeFrom: axisName === 'horizontal' ? 'paragraph' : 'line',
      referenceFrame: axisName === 'horizontal' ? 'paragraph' as const : 'line' as const,
      choiceKind: 'offset' as const, choiceValue: 0, baseStartPt: origin, baseEndPt: origin + 100,
      resolvedOriginPt: origin, pageParity: null,
    });
    const anchorFrame: AnchorFrameResult = {
      status: 'resolved', occurrenceId: 'frame-occurrence',
      axes: { horizontal: axis('horizontal', 4), vertical: axis('vertical', 5) }, issues: [],
      geometry: {
        objectFrame: rect(4, 5), inkBounds: rect(4, 5), wrapBounds: rect(3, 4, 12, 12),
        size: {
          horizontal: { source: 'extent', valuePt: 10, relativeFrom: null, referenceFrame: null, fraction: null },
          vertical: { source: 'extent', valuePt: 10, relativeFrom: null, referenceFrame: null, fraction: null },
        },
        parentEffectExtent: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
        wrap: {
          kind: 'square', side: 'bothSides',
          distances: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
          distanceSources: { top: 'implicit-zero', right: 'implicit-zero', bottom: 'implicit-zero', left: 'implicit-zero' },
          effectExtent: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 }, effectExtentSource: 'none',
          coordinateSpace: null, polygon: { edited: false, points: [{ xPt: 4, yPt: 5 }] },
        },
        transform: { coordinateSpace: 'anchor-frame', groupApplication: 'parser-resolved-child-frame', group: null },
      },
    };
    const localChild = paragraph('vertical-local-child');
    const verticalBox: TextBoxLayout = {
      kind: 'textbox', id: 'vertical-box', source: source([0, 4]), flowDomainId: 'textbox', ordinaryFlow: false,
      flowBounds: rect(8, 9), inkBounds: rect(8, 9), contentBounds: rect(9, 10), advancePt: 0,
      story: {
        story: 'textbox',
        flowBounds: localChild.flowBounds,
        inkBounds: localChild.inkBounds,
        blocks: [localChild],
        advancePt: localChild.advancePt,
        diagnostics: [],
      },
      transform: { a: 0, b: 1, c: -1, d: 0, e: 13, f: 14 },
      writingMode: 'vertical-rl', verticalMode: 'vert',
      insets: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    };
    const retained = {
      ...paragraph(),
      anchorFrames: [anchorFrame],
      anchorCollisions: [{
        occurrenceId: anchorFrame.occurrenceId,
        bounds: rect(4, 5),
        horizontalOwnership: 'host' as const,
        verticalOwnership: 'host' as const,
      }],
      textBoxes: [verticalBox],
      lineNumbers: [{ lineIndex: 0, counterValue: 12, bounds: rect(0, 2),
        paintOps: [{ kind: 'text' as const, text: '12', origin: { xPt: 0, yPt: 10 },
          font: '10pt serif', color: '#000000', textAlign: 'right' as const }] }],
    };

    const projected = projectBodyOccurrence(retained, options);

    expect(projected.anchorFrames![0]).toMatchObject({
      occurrenceId: 'page 2/header/anchor/frame-occurrence',
      axes: { horizontal: { resolvedOriginPt: 24 }, vertical: { resolvedOriginPt: 35 } },
      geometry: { objectFrame: { xPt: 24, yPt: 35 } },
    });
    expect(projected.anchorCollisions![0]).toEqual({
      occurrenceId: 'page 2/header/anchor/frame-occurrence',
      bounds: rect(24, 35),
      horizontalOwnership: 'host',
      verticalOwnership: 'host',
    });
    expect(projected.lineNumbers![0]).toMatchObject({ bounds: { xPt: 20, yPt: 32 },
      paintOps: [{ origin: { xPt: 20, yPt: 40 } }] });
    expect(projected.textBoxes[0]!.flowDomainId)
      .toBe('body/page-2/occurrence/page%202%2Fheader/textbox/vertical-box');
    expect(projected.textBoxes[0]!.story.blocks[0]!.flowBounds).toEqual(localChild.flowBounds);
  });
});
