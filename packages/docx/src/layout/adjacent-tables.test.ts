import { describe, expect, it } from 'vitest';
import {
  normalizeAdjacentTables,
  type AdjacentTableSequenceInput,
} from './adjacent-tables.js';
import type { BodyElement } from '../types.js';

function table(rowCount: number): BodyElement {
  return {
    type: 'table',
    colWidths: [10],
    rows: Array.from({ length: rowCount }, () => ({ cells: [] })),
    borders: { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left',
  } as unknown as BodyElement;
}

function paragraph(): BodyElement {
  return { type: 'paragraph' } as unknown as BodyElement;
}

/** Build a body-sequence entry from parser-owned §17.4.37 membership facts.
 * `sequenceId === null` models a table for which the parser preserved no
 * logical-table identity (a hand-built public table). */
function input(
  element: BodyElement,
  sequenceId: string | null,
  logicalRowOffset = 0,
  logicalTotalRows = element.type === 'table' ? element.rows.length : 0,
): AdjacentTableSequenceInput {
  if (element.type !== 'table' || sequenceId === null) {
    return { element, table: null };
  }
  return {
    element,
    table: {
      logicalSequenceId: sequenceId,
      logicalRowOffset,
      logicalTotalRows,
      rowCount: element.rows.length,
    },
  };
}

const collect = (body: readonly AdjacentTableSequenceInput[]) => [...normalizeAdjacentTables(body)];

describe('adjacent table normalization (ECMA-376 Part 1 §17.4.37)', () => {
  it('groups adjacent tables the parser assigned to one logical sequence', () => {
    const first = table(1);
    const second = table(2);

    expect(collect([
      input(first, 'seq:0', 0, 3),
      input(second, 'seq:0', 1, 3),
    ])).toEqual([{
      kind: 'adjacent-table-group',
      logicalSequenceId: 'seq:0',
      tables: [first, second],
    }]);
  });

  it('keeps parser-distinct logical sequences apart even when adjacent', () => {
    const first = table(1);
    const second = table(1);

    expect(collect([
      input(first, 'seq:0'),
      input(second, 'seq:7'),
    ])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: second },
    ]);
  });

  it('treats an intervening body element as a barrier', () => {
    const first = table(1);
    const between = paragraph();
    const second = table(1);

    expect(collect([
      input(first, 'seq:0'),
      input(between, null),
      input(second, 'seq:9'),
    ])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: between },
      { kind: 'body-element', element: second },
    ]);
  });

  it('keeps a table with no preserved parser identity standalone', () => {
    const first = table(1);
    const second = table(1);

    expect(collect([
      input(first, null),
      input(second, null),
    ])).toEqual([
      { kind: 'body-element', element: first },
      { kind: 'body-element', element: second },
    ]);
  });

  it('emits a single-member logical sequence as a plain body element', () => {
    const only = table(2);

    expect(collect([input(only, 'seq:0', 0, 2)])).toEqual([
      { kind: 'body-element', element: only },
    ]);
  });

  it('rejects a parser sequence whose row offsets are inconsistent', () => {
    const first = table(1);
    const second = table(2);

    expect(() => collect([
      input(first, 'seq:0', 0, 3),
      input(second, 'seq:0', 2, 3),
    ])).toThrow(/inconsistent/);
  });

  it('rejects a parser sequence whose member rows do not total', () => {
    const first = table(1);
    const second = table(1);

    expect(() => collect([
      input(first, 'seq:0', 0, 3),
      input(second, 'seq:0', 1, 3),
    ])).toThrow(/incomplete/);
  });
});
