import { expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const STYLES: Styles = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function worksheet(rightToLeft: boolean): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
    rightToLeft,
  } as Worksheet;
}

interface Segment { x1: number; y1: number; x2: number; y2: number; stroke: string }
interface Fill { x: number; y: number; w: number; h: number; fill: string }
interface TextPaint { text: string; fill: string }

function recordingCtx(width = 300, height = 120): {
  ctx: CanvasRenderingContext2D;
  segments: Segment[];
  fills: Fill[];
  texts: TextPaint[];
} {
  const segments: Segment[] = [];
  const fills: Fill[] = [];
  const texts: TextPaint[] = [];
  let strokeStyle = '#000';
  let fillStyle = '#000';
  let cursor: [number, number] | null = null;
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    font: '11px sans-serif',
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string) { fillStyle = value; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string) { strokeStyle = value; },
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 8 }),
    fillText: (text: string) => { texts.push({ text, fill: fillStyle }); },
    strokeText: () => {},
    fillRect: (x: number, y: number, w: number, h: number) => {
      fills.push({ x, y, w, h, fill: fillStyle });
    },
    strokeRect: () => {}, clearRect: () => {},
    beginPath: () => { cursor = null; }, closePath: () => {},
    moveTo: (x: number, y: number) => { cursor = [x, y]; },
    lineTo: (x: number, y: number) => {
      if (cursor) segments.push({ x1: cursor[0], y1: cursor[1], x2: x, y2: y, stroke: strokeStyle });
      cursor = [x, y];
    },
    rect: () => {}, arc: () => {}, fill: () => {}, stroke: () => {}, clip: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, setLineDash: () => {}, setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, segments, fills, texts };
}


import { viewportPreviewBlocker } from './internal/worksheet-preview-eligibility.js';

it('checks visible merged right border whose style comes from unloaded bottom corner', () => {
  const partial = worksheet(false);
  partial.mergeCells = [{top: 1, left: 1, bottom: 200, right: 2}];
  partial.rows = Array.from({length: 128}, (_, i) => ({index: i + 1, height: null, cells: []}));
  partial.rows[0].cells = [{row: 1, col: 1, value: {type: 'empty'}, styleIndex: 0}];
  const full = {...partial, rows: [...partial.rows, {index: 200, height: null, cells: [{row: 200, col: 2, value: {type: 'empty'}, styleIndex: 1}]}]} as Worksheet;
  const styles = {...STYLES, borders: [{}, {right: {style: 'thick', color: 'FF0000'}}], cellXfs: [STYLES.cellXfs[0], {...STYLES.cellXfs[0], borderId: 1}]} as Styles;
  const viewport = {row: 1, col: 1, rows: 30, cols: 12};
  const p = recordingCtx(900, 600), f = recordingCtx(900, 600);
  expect(viewportPreviewBlocker(partial, viewport, 128)).toBe('merge-range');
  renderViewport(p.ctx, partial, styles, viewport);
  renderViewport(f.ctx, full, styles, viewport);
  expect(p.segments).not.toEqual(f.segments);
});

it('checks CF in the frozen row', () => {
  const partial = worksheet(false);
  partial.freezeRows = 1;
  partial.rows = Array.from({length: 128}, (_, i) => ({index: i + 1, height: null, cells: []}));
  partial.rows[0].cells = [{row: 1, col: 1, value: {type: 'number', number: 5}, styleIndex: 0}];
  partial.conditionalFormats = [{sqref: [{top:1, left:1, bottom:1, right:1}], rules:[{type:'cellIs', operator:'greaterThan', formulas:['A500'], dxfId:0, priority:1}]}];
  const full = {...partial, rows: [...partial.rows, {index:500, height:null, cells:[{row:500, col:1, value:{type:'number', number:3}, styleIndex:0}]}]} as Worksheet;
  const styles = {...STYLES, dxfs: [{font:null, border:null, fill:{patternType:'solid', fgColor:'#FF0000', bgColor:'#FF0000'}}]} as Styles;
  const viewport = {row: 2, col: 1, rows: 29, cols: 12};
  expect(viewportPreviewBlocker(partial, { ...viewport, row: 1, rows: 30 }, 128)).toBe('conditional-format-range');
  const p=recordingCtx(900,600), f=recordingCtx(900,600);
  renderViewport(p.ctx, partial, styles, viewport, {freezeRows:1});
  renderViewport(f.ctx, full, styles, viewport, {freezeRows:1});
  expect(p.fills).not.toEqual(f.fills);
});
