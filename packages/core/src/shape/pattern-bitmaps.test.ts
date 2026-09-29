import { describe, expect, it } from 'vitest';
import { buildPatternBitmap } from './pattern-bitmaps.js';
import { resolveFill } from './paint.js';

// Independently transcribed from PowerPoint's 54 two-colour controls: each
// 16-digit string is eight row bytes, MSB left. The second mask records cells
// that are a 50/50 foreground/background blend in the exported PDF tile.
const OFFICE_MASKS: Record<string, [string, string]> = {
  pct5: ['8000000008000000', '0000000000000000'],
  pct10: ['8000080080000800', '0000000000000000'],
  pct20: ['8800220088002200', '0000000000000000'],
  pct25: ['8822882288228822', '0000000000000000'],
  pct30: ['AA44AA11AA44AA11', '0000000000000000'],
  pct40: ['AA55AA51AA55AA15', '0000000000000000'],
  pct50: ['AA55AA55AA55AA55', '0000000000000000'],
  pct60: ['EE55BB55EE55BB55', '0000000000000000'],
  pct70: ['77DD77DD77DD77DD', '0000000000000000'],
  pct75: ['77FFDDFF77FFDDFF', '0000000000000000'],
  pct80: ['EFFFFEFFEFFFFEFF', '0000000000000000'],
  pct90: ['FFFFFFF7FFFFFF7F', '0000000000000000'],
  horz: ['FF00000000000000', '0000000000000000'],
  vert: ['8080808080808080', '0000000000000000'],
  ltHorz: ['FF000000FF000000', '0000000000000000'],
  ltVert: ['8888888888888888', '0000000000000000'],
  dkHorz: ['FFFF0000FFFF0000', '0000000000000000'],
  dkVert: ['CCCCCCCCCCCCCCCC', '0000000000000000'],
  narHorz: ['FF00FF00FF00FF00', '0000000000000000'],
  narVert: ['5555555555555555', '0000000000000000'],
  dashHorz: ['F00000000F000000', '0000000000000000'],
  dashVert: ['8080808008080808', '0000000000000000'],
  cross: ['FF80808080808080', '0000000000000000'],
  dnDiag: ['8040201008040201', '41A05028140A0582'],
  upDiag: ['0102040810204080', '82050A142850A041'],
  ltDnDiag: ['8844221188442211', '0000000000000000'],
  ltUpDiag: ['1122448811224488', '0000000000000000'],
  dkDnDiag: ['CC663399CC663399', '0000000000000000'],
  dkUpDiag: ['3366CC993366CC99', '0000000000000000'],
  wdDnDiag: ['C1E070381C0E0783', '0000000000000000'],
  wdUpDiag: ['83070E1C3870E0C1', '0000000000000000'],
  dashDnDiag: ['0000884422110000', '0000000000000000'],
  dashUpDiag: ['0000112244880000', '0000000000000000'],
  diagCross: ['8142241818244281', '42A55A24245AA542'],
  smCheck: ['9966669999666699', '0000000000000000'],
  lgCheck: ['F0F0F0F00F0F0F0F', '0000000000000000'],
  smGrid: ['FF888888FF888888', '0000000000000000'],
  lgGrid: ['FF80808080808080', '0000000000000000'],
  dotGrid: ['AA00800080008000', '0000000000000000'],
  smConfetti: ['8008400210012004', '0000000000000000'],
  lgConfetti: ['B130031BD8C00C8D', '0000000000000000'],
  horzBrick: ['FF808080FF080808', '0000000000000000'],
  diagBrick: ['0102040818244281', '0000000000000000'],
  solidDmnd: ['10387CFE7C381000', '0000000000000000'],
  openDmnd: ['8244281028448201', '0000000000000000'],
  dotDmnd: ['8000220008002200', '0000000000000000'],
  plaid: ['AA55AA55F0F0F0F0', '0000000000000000'],
  sphere: ['77898F8F7798F8F8', '0000000000000000'],
  weave: ['8854224588142251', '0000000000000000'],
  divot: ['0010081000800180', '0000000000000000'],
  shingle: ['038448300C020101', '0000000000000000'],
  wave: ['001825C0001825C0', '0000000000000000'],
  trellis: ['FF66FF99FF66FF99', '0000000000000000'],
  zigZag: ['8142241881422418', '0000000000000000'],
};

class RecordingCanvas {
  readonly width: number;
  readonly height: number;
  readonly pixels: string[][];
  constructor(width: number, height: number) {
    this.width = width;
    this.height = height;
    this.pixels = Array.from({ length: height }, () => Array(width).fill(''));
  }
  getContext() {
    const pixels = this.pixels;
    return {
      fillStyle: '',
      fillRect(this: { fillStyle: string }, x: number, y: number, w: number, h: number) {
        for (let row = y; row < y + h; row++) {
          for (let col = x; col < x + w; col++) pixels[row][col] = this.fillStyle;
        }
      },
    };
  }
}

describe('DrawingML preset pattern bitmap (PowerPoint PDF oracle)', () => {
  it('paints every specified preset, including half-covered diagonal cells', () => {
    const previous = globalThis.OffscreenCanvas;
    Object.defineProperty(globalThis, 'OffscreenCanvas', {
      configurable: true, value: RecordingCanvas,
    });
    try {
      expect(Object.keys(OFFICE_MASKS)).toHaveLength(54);
      for (const [preset, [full, half]] of Object.entries(OFFICE_MASKS)) {
        const canvas = buildPatternBitmap(preset, 'D21D54', '12CED4') as RecordingCanvas | null;
        expect(canvas, preset).not.toBeNull();
        expect([canvas!.width, canvas!.height], preset).toEqual([64, 64]);
        const actual = Array.from({ length: 8 }, (_, y) => Array.from({ length: 8 }, (_, x) => {
          const color = canvas!.pixels[y * 8 + 4][x * 8 + 4];
          if (color === 'rgba(210,29,84,1)') return 2;
          if (color === 'rgba(114,118,148,1)') return 1;
          if (color === 'rgba(18,206,212,1)') return 0;
          throw new Error(`${preset}: unexpected colour ${color}`);
        }));
        for (let y = 0; y < 8; y++) {
          const fullByte = Number.parseInt(full.slice(y * 2, y * 2 + 2), 16);
          const halfByte = Number.parseInt(half.slice(y * 2, y * 2 + 2), 16);
          for (let x = 0; x < 8; x++) {
            const mask = 1 << (7 - x);
            expect(actual[y][x], `${preset} (${x}, ${y})`).toBe(
              fullByte & mask ? 2 : halfByte & mask ? 1 : 0,
            );
          }
        }
      }
    } finally {
      Object.defineProperty(globalThis, 'OffscreenCanvas', {
        configurable: true, value: previous,
      });
    }
  });

  it('sizes the 8 pt tile in a CSS pixel coordinate system with slide-origin phase', () => {
    const previous = globalThis.OffscreenCanvas;
    Object.defineProperty(globalThis, 'OffscreenCanvas', {
      configurable: true, value: RecordingCanvas,
    });
    try {
      let matrix: DOMMatrix2DInit | undefined;
      const pattern = { setTransform(value: DOMMatrix2DInit) { matrix = value; } };
      const ctx = { createPattern: () => pattern } as unknown as CanvasRenderingContext2D;
      const result = resolveFill(
        { fillType: 'pattern', preset: 'pct30', fg: 'D21D54', bg: '12CED4' },
        ctx, 102.24, 30.24, 390.24, 95.76,
      );
      expect(result).toBe(pattern);
      expect(matrix).toMatchObject({ a: 1 / 6, d: 1 / 6, e: 0, f: 0 });
      const shifted = resolveFill(
        { fillType: 'pattern', preset: 'pct30', fg: 'D21D54', bg: '12CED4' },
        ctx, 514.8, 154.08, 390.24, 95.76,
      );
      expect(shifted).toBe(pattern);
      expect(matrix).toMatchObject({ a: 1 / 6, d: 1 / 6, e: 0, f: 0 });

      // DOCX paints in point units, so its same 8 pt tile needs unit scale 1.
      const pointPattern = { setTransform(value: DOMMatrix2DInit) { matrix = value; } };
      const pointCtx = { createPattern: () => pointPattern } as unknown as CanvasRenderingContext2D;
      resolveFill(
        { fillType: 'pattern', preset: 'pct30', fg: 'D21D54', bg: '12CED4' },
        pointCtx, 90, 62.18, 432, 72, 0, 1,
      );
      expect(matrix).toMatchObject({ a: 1 / 8, d: 1 / 8, e: 0, f: 0 });
    } finally {
      Object.defineProperty(globalThis, 'OffscreenCanvas', {
        configurable: true, value: previous,
      });
    }
  });
});
