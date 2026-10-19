import { describe, it, expect, beforeEach } from 'vitest';
import { drawUnderline } from './underline.js';

// A minimal 2D-context stub that records the drawing calls drawUnderline makes,
// so we can assert the geometry / dash dispatch for each DrawingML
// ST_TextUnderlineType (§20.1.10.82) style without a real canvas.
interface Op {
  op: string;
  args: number[];
}
function makeCtx() {
  const ops: Op[] = [];
  let dash: number[] = [];
  const dashHistory: number[][] = [];
  const ctx = {
    strokeStyle: '',
    lineWidth: 0,
    setLineDash(d: number[]) {
      dash = d.slice();
      dashHistory.push(d.slice());
    },
    getLineDash() {
      return dash;
    },
    beginPath() {
      ops.push({ op: 'beginPath', args: [] });
    },
    moveTo(x: number, y: number) {
      ops.push({ op: 'moveTo', args: [x, y] });
    },
    lineTo(x: number, y: number) {
      ops.push({ op: 'lineTo', args: [x, y] });
    },
    stroke() {
      ops.push({ op: 'stroke', args: [] });
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, ops, dashHistory };
}

describe('drawUnderline (core, DrawingML ST_TextUnderlineType §20.1.10.82)', () => {
  let h: ReturnType<typeof makeCtx>;
  beforeEach(() => {
    h = makeCtx();
  });

  it('default (undefined) draws a single solid rule with an empty dash', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', undefined);
    // One straight segment: moveTo + lineTo + stroke.
    const move = h.ops.find((o) => o.op === 'moveTo');
    const line = h.ops.find((o) => o.op === 'lineTo');
    expect(move).toBeDefined();
    expect(line).toBeDefined();
    // Spans x=10 → x=60 (x + width).
    expect(move!.args[0]).toBe(10);
    expect(line!.args[0]).toBe(60);
    // The final dash applied for a solid rule is empty.
    expect(h.dashHistory.at(-1)).toEqual([]);
    // Single rule: exactly one moveTo/lineTo pair.
    expect(h.ops.filter((o) => o.op === 'moveTo')).toHaveLength(1);
    expect(h.ops.filter((o) => o.op === 'lineTo')).toHaveLength(1);
  });

  it('sng behaves identically to the default single rule', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'sng');
    expect(h.ops.filter((o) => o.op === 'moveTo')).toHaveLength(1);
    expect(h.dashHistory.at(-1)).toEqual([]);
  });

  it('dbl draws two parallel solid rules', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'dbl');
    // Two moveTo + two lineTo (both rails), one stroke.
    expect(h.ops.filter((o) => o.op === 'moveTo')).toHaveLength(2);
    expect(h.ops.filter((o) => o.op === 'lineTo')).toHaveLength(2);
    const [m1, m2] = h.ops.filter((o) => o.op === 'moveTo');
    // The two rails sit at distinct y offsets straddling the rule y.
    expect(m1.args[1]).not.toBe(m2.args[1]);
  });

  it('dotted applies a non-empty dash pattern before stroking', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'dotted');
    // The dash set immediately before the straight stroke is non-empty.
    const applied = h.dashHistory.find((d) => d.length > 0);
    expect(applied).toBeDefined();
    // A single straight rule (dotted is drawn as a dashed straight line).
    expect(h.ops.filter((o) => o.op === 'moveTo')).toHaveLength(1);
  });

  it('dash applies a dash pattern', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'dash');
    expect(h.dashHistory.some((d) => d.length > 0)).toBe(true);
  });

  it('dotDash applies a 4-element dash pattern', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'dotDash');
    const applied = h.dashHistory.find((d) => d.length === 4);
    expect(applied).toBeDefined();
  });

  it('dotDotDash applies a 6-element dash pattern', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#000000', 'dotDotDash');
    const applied = h.dashHistory.find((d) => d.length === 6);
    expect(applied).toBeDefined();
  });

  it('wavy traces a multi-segment polyline (many lineTo)', () => {
    drawUnderline(h.ctx, 0, 100, 60, 20, '#000000', 'wavy');
    // Sine polyline: far more than the single lineTo of a straight rule.
    expect(h.ops.filter((o) => o.op === 'lineTo').length).toBeGreaterThan(5);
    // Wavy omits setLineDash mutation for the wave itself (only the initial []).
    expect(h.ops.filter((o) => o.op === 'stroke')).toHaveLength(1);
  });

  it('wavyDbl traces two wave polylines (two strokes)', () => {
    drawUnderline(h.ctx, 0, 100, 60, 20, '#000000', 'wavyDbl');
    expect(h.ops.filter((o) => o.op === 'stroke')).toHaveLength(2);
  });

  it('*Heavy variants thicken the line weight ~1.8×', () => {
    const plain = makeCtx();
    const heavy = makeCtx();
    drawUnderline(plain.ctx, 10, 100, 50, 20, '#000', 'dash');
    drawUnderline(heavy.ctx, 10, 100, 50, 20, '#000', 'dashHeavy');
    // lineWidth is the last-written value; heavy is 1.8× the plain weight.
    expect((heavy.ctx as unknown as { lineWidth: number }).lineWidth).toBeCloseTo(
      (plain.ctx as unknown as { lineWidth: number }).lineWidth * 1.8,
      6,
    );
  });

  it('bare "heavy" (§20.1.10.82 ST_TextUnderlineType, pptx <a:u val="heavy">) thickens the line weight ~1.8×', () => {
    const plain = makeCtx();
    const heavy = makeCtx();
    drawUnderline(plain.ctx, 10, 100, 50, 20, '#000', 'sng');
    drawUnderline(heavy.ctx, 10, 100, 50, 20, '#000', 'heavy');
    expect((heavy.ctx as unknown as { lineWidth: number }).lineWidth).toBeCloseTo(
      (plain.ctx as unknown as { lineWidth: number }).lineWidth * 1.8,
      6,
    );
  });

  it('uses the provided colour for strokeStyle', () => {
    drawUnderline(h.ctx, 10, 100, 50, 20, '#ff0000', 'sng');
    expect((h.ctx as unknown as { strokeStyle: string }).strokeStyle).toBe('#ff0000');
  });
});

describe('drawUnderline with the face rule (#1896 Windows PowerPoint controls)', () => {
  // Arial at 48 pt: post underlinePosition -217, underlineThickness 150 / 2048.
  const rule = { topPx: 217 / 2048 * 48, thicknessPx: 150 / 2048 * 48 };

  it('centres the single rule on the face underline', () => {
    const { ctx, ops } = makeCtx();
    drawUnderline(ctx, 0, 100, 50, 48, '#000', 'sng', 1, undefined, rule);
    const move = ops.find((op) => op.op === 'moveTo')!;
    expect(move.args[1]).toBeCloseTo(100 + rule.topPx + rule.thicknessPx / 2, 0);
    expect(ctx.lineWidth).toBeCloseTo(rule.thicknessPx, 6);
  });

  it('draws heavy 1.5x as thick with its bottom at the single rule centre', () => {
    const { ctx, ops } = makeCtx();
    drawUnderline(ctx, 0, 100, 50, 48, '#000', 'heavy', 1, undefined, rule);
    const move = ops.find((op) => op.op === 'moveTo')!;
    const singleCentre = 100 + rule.topPx + rule.thicknessPx / 2;
    expect(ctx.lineWidth).toBeCloseTo(rule.thicknessPx * 1.5, 6);
    expect(move.args[1] + ctx.lineWidth / 2).toBeCloseTo(singleCentre, 0);
  });
});
