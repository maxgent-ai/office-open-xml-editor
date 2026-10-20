import { describe, expect, it } from 'vitest';
import { compileCf, evaluateCf } from './conditional-format.js';
import type {
  Cell, CellValue, CfRule, CfValue, ConditionalFormat, Dxf, Worksheet, WorksheetCellRange,
} from './types.js';

// Shared CF evaluator used by every XLSX consumer (main renderer, render
// workers). The Node facade has no XLSX render API, so this is the entry.
// Library policy (#1547): a rule the evaluator cannot evaluate neither
// formats the cell nor stops lower rules.

const RED: Dxf = { font: null, fill: { patternType: 'solid', fgColor: '#FF0000', bgColor: '#FF0000' }, border: null };
const BLUE: Dxf = {
  font: { bold: false, italic: false, underline: false, strike: false, size: 11, color: '#0000FF', name: null },
  fill: null, border: null,
};
const DXFS = [RED, BLUE];

const at = (row: number, col: number, value: CellValue): Cell => ({ row, col, value, styleIndex: 0 });
const num = (row: number, col: number, n: number): Cell => at(row, col, { type: 'number', number: n });
const err = (row: number, col: number, e: string): Cell => at(row, col, { type: 'error', error: e });

function sheet(cells: Cell[], conditionalFormats: ConditionalFormat[]): Worksheet {
  const rows = new Map<number, Cell[]>();
  for (const c of cells) rows.set(c.row, [...(rows.get(c.row) ?? []), c]);
  return {
    name: 'Sheet1',
    rows: [...rows].sort(([a], [b]) => a - b).map(([index, rowCells]) => ({ index, height: null, cells: rowCells })),
    colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats, images: [], charts: [],
  };
}

const C1: WorksheetCellRange = { top: 1, left: 3, bottom: 1, right: 3 };
const A1_A3: WorksheetCellRange = { top: 1, left: 1, bottom: 3, right: 1 };
const lowerRed: CfRule = { type: 'expression', formula: 'TRUE', dxfId: 0, priority: 9, stopIfTrue: false };

function evaluate(ws: Worksheet, row: number, col: number) {
  const target = ws.rows.find((r) => r.index === row)?.cells.find((c) => c.col === col);
  return evaluateCf(target, row, col, compileCf(ws), DXFS);
}

describe('parser-preserved unimplemented formula inlets', () => {
  const marker = (mask: number): CfRule => ({ type: 'other', kind: 'expression',
    priority: 1, stopIfTrue: true, unsupportedFormulaPhases: mask } as CfRule);
  it('marker-only blocks do not rescan a warm worksheet cell index', () => {
    const ws = sheet([], [{sqref:[C1],rules:[marker(1)]}]);
    Object.defineProperty(ws, 'rows', {get: () => {throw new Error('unexpected cached-cell rescan');}});
    const compiled = compileCf(ws, new Map());
    expect(evaluateCf(undefined, 1, 3, compiled, DXFS)).toEqual({});
  });
  it('neither formats nor stops lower rules', () => {
    const ws = sheet([num(1, 3, 5)], [{ sqref: [C1], rules: [marker(15), lowerRed] }]);
    expect(evaluate(ws, 1, 3).fill?.fgColor).toBe('#FF0000');
  });
});

describe('expression ingress', () => {
  // Mathematically TRUE (1*1+2*1>0), but SUMPRODUCT is not in the
  // evaluator's function table.
  const sumproductCells = [num(1, 1, 1), num(2, 1, 2), num(1, 2, 1), num(2, 2, 1), num(1, 3, 5)];
  const sumproduct: CfRule = {
    type: 'expression', formula: 'SUMPRODUCT(A1:A2,B1:B2)>0', dxfId: 1, priority: 1, stopIfTrue: true,
  };

  it('an unsupported SUMPRODUCT keeps the lower rule and does not stop', () => {
    const ws = sheet(sumproductCells, [{ sqref: [C1], rules: [lowerRed, sumproduct] }]);
    const before = structuredClone(ws.rows);
    const result = evaluate(ws, 1, 3);
    expect(result.fontColor).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
    // Cached values and formulas are never recalculated or mutated.
    expect(ws.rows).toEqual(before);
  });

  it('valid FALSE / zero / Excel error results are no-matches that do not stop', () => {
    for (const formula of ['FALSE', '0', '#N/A=1', '1/0>0', 'A1>100']) {
      const rule: CfRule = { type: 'expression', formula, dxfId: 1, priority: 1, stopIfTrue: true };
      const result = evaluate(sheet(sumproductCells, [{ sqref: [C1], rules: [rule, lowerRed] }]), 1, 3);
      expect(result.fill?.fgColor, formula).toBe('#FF0000');
    }
  });

  it('invalid syntax neither formats nor stops', () => {
    const rule: CfRule = { type: 'expression', formula: '1+', dxfId: 1, priority: 1, stopIfTrue: true };
    const result = evaluate(sheet([num(1, 3, 5)], [{ sqref: [C1], rules: [rule, lowerRed] }]), 1, 3);
    expect(result.fontColor).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });

  it('a supported stopping rule prevents evaluation of lower rules', () => {
    const stop: CfRule = { type: 'expression', formula: 'TRUE', dxfId: 1, priority: 1, stopIfTrue: true };
    const lower = { ...sumproduct, priority: 2, stopIfTrue: false };
    const result = evaluate(sheet(sumproductCells, [{ sqref: [C1], rules: [stop, lower, lowerRed] }]), 1, 3);
    expect(result.fontColor).toBe('#0000FF');
    expect(result.fill).toBeUndefined();
  });

  it('data-dependent limitations are decided at runtime, not statically', () => {
    // §18.17.7.278 does not establish fractional number-digits: unsupported
    // only when the cached cell actually holds a fraction.
    const rule: CfRule = { type: 'expression', formula: 'ROUND(1.25,$D$1)=1.3', dxfId: 1, priority: 1, stopIfTrue: false };
    expect(evaluate(sheet([num(1, 3, 5), num(1, 4, 1)], [{ sqref: [C1], rules: [rule] }]), 1, 3).fontColor)
      .toBe('#0000FF');
    expect(evaluate(sheet([num(1, 3, 5), num(1, 4, 1.5)], [{ sqref: [C1], rules: [rule] }]), 1, 3).fontColor)
      .toBeUndefined();
  });
});

describe('cellIs ingress', () => {
  const cellIs = (operator: string, formulas: string[]): CfRule =>
    ({ type: 'cellIs', operator, formulas, dxfId: 1, priority: 1, stopIfTrue: true });
  const run = (rule: CfRule, extra: Cell[] = []) =>
    evaluate(sheet([num(1, 3, 5), ...extra], [{ sqref: [C1], rules: [rule, lowerRed] }]), 1, 3);

  it('an operand outside the literal/reference grammar neither formats nor stops', () => {
    const result = run(cellIs('greaterThan', ['A1+1']), [num(1, 1, 1)]);
    expect(result.fill?.fgColor).toBe('#FF0000');
    expect(result.fontColor).toBeUndefined();
  });

  it('a cached Excel error, a blank and a type mismatch are no-matches', () => {
    for (const extra of [[err(1, 2, '#N/A')], [], [at(1, 2, { type: 'bool', bool: false })], [at(1, 2, { type: 'text', text: 'x' })]]) {
      const result = run(cellIs('greaterThan', ['$B$1']), extra);
      expect(result.fill?.fgColor).toBe('#FF0000');
      expect(result.fontColor).toBeUndefined();
    }
  });

  it('literal and reference controls are supported', () => {
    expect(run(cellIs('greaterThan', ['4'])).fontColor).toBe('#0000FF');
    expect(run(cellIs('greaterThan', ['$B$1']), [num(1, 2, 3)]).fontColor).toBe('#0000FF');
  });

  it('a missing operand or an unknown operator neither formats nor stops', () => {
    for (const rule of [cellIs('between', ['1']), cellIs('greaterThan', []), cellIs('greaterish', ['1'])]) {
      const result = run(rule);
      expect(result.fontColor).toBeUndefined();
      expect(result.fill?.fgColor).toBe('#FF0000');
    }
  });
});

describe('scale-rule activity ingress', () => {
  const bar = (activeFormula?: string, extra: Partial<Extract<CfRule, { type: 'dataBar' }>> = {}): CfRule => ({
    type: 'dataBar', color: '#638EC6', min: { kind: 'min', value: null }, max: { kind: 'max', value: null },
    priority: 1, gradient: true, activeFormula, stopIfTrue: true, ...extra,
  });
  const run = (rule: CfRule, extra: Cell[] = []) =>
    evaluate(sheet([num(1, 1, 1), num(2, 1, 5), num(3, 1, 9), ...extra], [{ sqref: [A1_A3], rules: [rule, lowerRed] }]), 2, 1);

  it('an unsupported activity formula neither paints nor stops', () => {
    for (const formula of ['SUM(A1:A2)', 'SUMPRODUCT(A1:A3)']) {
      const result = run(bar(formula));
      expect(result.dataBar, formula).toBeUndefined();
      expect(result.fill?.fgColor, formula).toBe('#FF0000');
    }
  });

  it('absent is active, valid zero is inactive, cached error is inactive', () => {
    expect(run(bar()).dataBar).toBeDefined();
    expect(run(bar('0')).dataBar).toBeUndefined();
    expect(run(bar('$B$1'), [err(1, 2, '#DIV/0!')]).dataBar).toBeUndefined();
  });
});

describe('cfvo threshold ingress', () => {
  const cells = [num(1, 1, 1), num(2, 1, 5), num(3, 1, 9)];
  const run = (rule: CfRule) => evaluate(sheet(cells, [{ sqref: [A1_A3], rules: [rule, lowerRed] }]), 2, 1);
  it('does not convert automatic or unknown threshold kinds into zero', () => {
    for (const kind of ['autoMin', 'autoMax', 'futureKind']) {
      const result = run({type:'dataBar', priority:1, stopIfTrue:true,
        color:'#638EC6', gradient:true, min:{kind,value:null}, max:{kind:'max',value:null}});
      expect(result.dataBar).toBeUndefined();
      expect(result.fill?.fgColor).toBe('#FF0000');
    }
  });
  it('a parser-marked x14 literal formula is not a standard literal threshold', () => {
    const result = run({type:'dataBar', priority:1, stopIfTrue:true,
      color:'#638EC6', gradient:true, extThresholdFormula:true,
      min:{kind:'formula',value:null},max:{kind:'num',value:'10'}});
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });
  const scale = (first: CfValue): CfRule => ({
    type: 'colorScale', priority: 1, stopIfTrue: true,
    stops: [{ ...first, color: '#000000' }, { kind: 'max', value: null, color: '#FFFFFF' }],
  });
  const bar = (min: CfValue): CfRule => ({
    type: 'dataBar', color: '#638EC6', min, max: { kind: 'max', value: null }, priority: 1, gradient: true, stopIfTrue: true,
  });
  const icons = (middle: CfValue): CfRule => ({
    type: 'iconSet', iconSet: '3Arrows', reverse: false, priority: 1, stopIfTrue: true,
    cfvos: [{ kind: 'percent', value: '0' }, middle, { kind: 'percent', value: '67' }],
  });

  it('formula thresholds (including the numeric-prefix 1+1) are skipped in all three types', () => {
    for (const value of ['$A$1', '1+1', 'MAX(A1:A3)']) {
      const formula: CfValue = { kind: 'formula', value };
      expect(run(scale(formula)).fill?.fgColor, value).toBe('#FF0000'); // no guessed scale; lower rule paints
      const dataBar = run(bar(formula));
      expect(dataBar.dataBar, value).toBeUndefined();
      expect(dataBar.fill?.fgColor, value).toBe('#FF0000');
      const iconSet = run(icons(formula));
      expect(iconSet.iconSet, value).toBeUndefined();
      expect(iconSet.fill?.fgColor, value).toBe('#FF0000');
    }
    // A num threshold written as a formula is not read as its numeric prefix.
    expect(run(bar({ kind: 'num', value: '1+1' })).dataBar).toBeUndefined();
  });

  it('a formula threshold without text is skipped', () => {
    const result = run(bar({ kind: 'formula', value: null }));
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });

  it('min / max / num / percent / literal formula controls stay valid and paint', () => {
    expect(run(scale({ kind: 'min', value: null })).fill?.fgColor).not.toBe('#FF0000');
    expect(run(bar({ kind: 'num', value: '2' })).dataBar).toBeDefined();
    expect(run(icons({ kind: 'percent', value: '33' })).iconSet).toBeDefined();
    expect(run(icons({ kind: 'formula', value: '5' })).iconSet).toBeDefined();
  });

  it('a linked x14 threshold formula marker skips the rule', () => {
    const rule: CfRule = { ...(bar({ kind: 'min', value: null }) as Extract<CfRule, { type: 'dataBar' }>), extThresholdFormula: true };
    const result = run(rule);
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });

  it('an inactive rule with an unresolvable threshold neither paints nor stops', () => {
    const rule: CfRule = { ...(bar({ kind: 'formula', value: '$A$1' }) as Extract<CfRule, { type: 'dataBar' }>), activeFormula: '0' };
    const result = run(rule);
    expect(result.dataBar).toBeUndefined();
    expect(result.fill?.fgColor).toBe('#FF0000');
  });
});
