import { readFileSync } from 'node:fs';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import type { Worksheet } from '@silurus/ooxml-xlsx';
import type { ModelSource, ModelSourceConfig } from '@silurus/ooxml-core';
import { DEFAULT_XLSX_WORKSHEET_POLICY } from '@silurus/ooxml-core/worker';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
// Test-only internal Node seams (no public API is added by this file).
import { loadSkiaForTests } from './test-imports.ts';
import { formatCellValue } from '../../xlsx/src/number-format.ts';
import { initSync, XlsxArchive } from '../../xlsx/src/wasm/xlsx_parser.js';
import { getWorksheetPolicy } from '../../xlsx/src/worksheet-policy-context.ts';
import { renderWorksheetViewport } from '../../xlsx/src/render-orchestrator.ts';
import { XlsxFindController, type FindCell } from '../../xlsx/src/find.ts';
import { CopyController } from '../../xlsx/src/internal/viewer/copy-controller.ts';
import { selectionStateFromReference, type XlsxSelectionState } from '../../xlsx/src/selection.ts';
import { materializeXlsxWorksheet } from './xlsx.ts';

type MaterializeOptions = NonNullable<Parameters<typeof materializeXlsxWorksheet>[2]>;
type RenderStyles = Parameters<typeof renderWorksheetViewport>[0]['styles'];

const LARGE_TIMEOUT = 120_000;
const DEFAULTS = DEFAULT_XLSX_WORKSHEET_POLICY.worksheet;

let directory = '';
let small: Uint8Array;

beforeAll(async () => {
  initSync({ module: readFileSync(new URL('../../xlsx/src/wasm/xlsx_parser_bg.wasm', import.meta.url)) });
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-finite-integration-'));
  const fixture = join(directory, 'small.xlsx');
  await generateSyntheticXlsx(fixture, { rows: 3, columns: 2 });
  small = new Uint8Array(await readFile(fixture));
});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

// ---------------------------------------------------------------- native helpers

function limits(maxRows: bigint): [bigint, bigint, bigint, bigint] {
  return [maxRows, BigInt(DEFAULTS.maxCells), BigInt(DEFAULTS.maxOwnedUtf8Bytes), BigInt(DEFAULTS.maxJsonBytes)];
}

function decodeJson(bytes: Uint8Array): unknown {
  return JSON.parse(new TextDecoder().decode(bytes));
}

function firstSheetName(archive: XlsxArchive): string {
  const index = decodeJson(archive.parse());
  if (typeof index === 'object' && index !== null && 'workbook' in index && typeof index.workbook === 'object' && index.workbook !== null && 'sheets' in index.workbook && Array.isArray(index.workbook.sheets)) {
    const first: unknown = index.workbook.sheets[0];
    if (typeof first === 'object' && first !== null && 'name' in first && typeof first.name === 'string') {
      return first.name;
    }
  }
  throw new Error('workbook index has no first sheet name');
}

/** Drain worksheet 0 through the real resumable cursor; returns the pull count. */
function drainFirstSheet(archive: XlsxArchive): number {
  const name = firstSheetName(archive);
  archive.open_sheet_cursor(0, name);
  try {
    for (let pulls = 1; pulls <= 10_000; pulls++) {
      archive.pull_sheet_cursor(1_024);
      if (archive.sheet_cursor_pull_finished()) {
        archive.acknowledge_sheet_cursor_terminal();
        return pulls;
      }
    }
    throw new Error('cursor never reached its terminal product');
  } finally {
    archive.close_sheet_cursor();
  }
}

function withArchive<T>(bytes: Uint8Array, use: (archive: XlsxArchive) => T): T {
  const archive = new XlsxArchive(new Uint8Array(bytes));
  try {
    return use(archive);
  } finally {
    archive.free();
  }
}

describe('native XlsxArchive worksheet limits (real generated WASM)', () => {
  it('accepts one setter before parse, refuses a second setter and keeps the archive usable', () => {
    withArchive(small, (archive) => {
      archive.set_worksheet_limits(...limits(3n));
      expect(() => archive.set_worksheet_limits(...limits(1n))).toThrow();
      // Setter rejection does not prevent the existing non-retaining cursor.
      expect(drainFirstSheet(archive)).toBeGreaterThan(0);
    });
  });

  it('refuses the setter after the archive has been parsed', () => {
    withArchive(small, (archive) => {
      archive.parse();
      expect(() => archive.set_worksheet_limits(...limits(3n))).toThrow();
    });
  });

  it('refuses 0 and MAX_SAFE_INTEGER without consuming the setter; a valid setter is then accepted', () => {
    withArchive(small, (archive) => {
      expect(() => archive.set_worksheet_limits(...limits(0n))).toThrow();
      expect(() => archive.set_worksheet_limits(...limits(BigInt(Number.MAX_SAFE_INTEGER)))).toThrow();
      archive.set_worksheet_limits(...limits(1n));
      expect(drainFirstSheet(archive)).toBeGreaterThan(0);
    });
  });

  it('configures sibling archives independently without changing streaming credits', () => {
    const low = new XlsxArchive(new Uint8Array(small));
    const high = new XlsxArchive(new Uint8Array(small));
    try {
      low.set_worksheet_limits(...limits(1n));
      high.set_worksheet_limits(...limits(3n));
      // The non-retaining native cursor keeps its existing streaming credits.
      expect(drainFirstSheet(low)).toBeGreaterThan(0);
      expect(drainFirstSheet(high)).toBeGreaterThan(0);
    } finally {
      low.free();
      high.free();
    }
  });
});

// ---------------------------------------------------------------- application ModelSource

function fakeSource(config: ModelSourceConfig = { minimal: true }) {
  const probe = new ArrayBuffer(24);
  const release = vi.fn();
  const source: ModelSource<'xlsx'> = {
    target: 'xlsx',
    claim: vi.fn(() => true),
    beginLoad: () => ({
      module: {
        protocol: 'ooxml-model-source-module/v1',
        target: 'xlsx',
        moduleUrl: new URL('./test-fixtures/xlsx-model-source.mjs', import.meta.url).href,
        config,
      },
      transfer: [probe],
      release,
    }),
  };
  const counters = new Float64Array(probe);
  return { source, release, closes: () => counters[0] };
}

describe('application minimal ModelSource without the worksheet-limit setter', () => {
  it('rejects maxRows 1 and materializes under maxRows 3, releasing and closing each load once', async () => {
    const low = fakeSource();
    await expect(materializeXlsxWorksheet(small, 0, {
      xlsxWorksheetLimits: { maxRows: 1 },
      modelSources: [low.source],
    })).rejects.toMatchObject({
      code: 'ooxml-resource-limit',
      details: expect.objectContaining({
        violation: expect.objectContaining({ configurable: true, limit: 1 }),
      }),
    });
    await vi.waitFor(() => expect(low.closes()).toBe(1));
    expect(low.release).toHaveBeenCalledTimes(1);

    const high = fakeSource();
    const worksheet = await materializeXlsxWorksheet(small, 0, {
      xlsxWorksheetLimits: { maxRows: 3 },
      modelSources: [high.source],
    });
    expect(worksheet.rows).toHaveLength(3);
    expect(high.source.claim).toHaveBeenCalled();
    await vi.waitFor(() => expect(high.closes()).toBe(1));
    expect(high.release).toHaveBeenCalledTimes(1);
  });

  it('rejects malformed limits before any claim or load', async () => {
    const fake = fakeSource();
    const options: MaterializeOptions = { xlsxWorksheetLimits: { maxRows: 0 }, modelSources: [fake.source] };
    await expect(materializeXlsxWorksheet(small, 0, options)).rejects.toBeInstanceOf(RangeError);
    expect(fake.source.claim).not.toHaveBeenCalled();
    expect(fake.release).not.toHaveBeenCalled();
    expect(fake.closes()).toBe(0);
  });
});

// ---------------------------------------------------------------- optional Canvas (Skia)

const skia = await loadSkiaForTests();

function indexStyles(bytes: Uint8Array, worksheet: Worksheet): RenderStyles {
  const p = getWorksheetPolicy(worksheet).worksheet;
  return withArchive(bytes, (archive) => {
    archive.set_worksheet_limits(
      BigInt(p.maxRows), BigInt(p.maxCells), BigInt(p.maxOwnedUtf8Bytes), BigInt(p.maxJsonBytes),
    );
    const index = decodeJson(archive.parse());
    if (typeof index !== 'object' || index === null || !('styles' in index)) throw new Error('index has no styles');
    return index.styles as RenderStyles;
  });
}

interface LargeCase {
  label: string;
  rows: number;
  columns: number;
  overrides: NonNullable<MaterializeOptions['xlsxWorksheetLimits']>;
  tailRef: string;
  tailText: string;
  tailViewport: { row: number; col: number; rows: number; cols: number };
}

const cases: LargeCase[] = [
  {
    label: '100001 rows x 1 column', rows: 100_001, columns: 1, overrides: { maxRows: 100_001 },
    tailRef: 'A100001', tailText: 'r100001c1', tailViewport: { row: 99_990, col: 1, rows: 12, cols: 3 },
  },
  {
    label: '12550 rows x 20 columns (251000 cells)', rows: 12_550, columns: 20, overrides: { maxCells: 251_000 },
    tailRef: 'T12550', tailText: '251020', tailViewport: { row: 12_540, col: 15, rows: 11, cols: 6 },
  },
];

describe.skipIf(!skia)('large finite worksheets through Canvas, find and copy seams', () => {
  for (const testCase of cases) {
    it(`renders, finds and copies the tail of ${testCase.label}`, async () => {
      if (!skia) return;
      const path = join(directory, `large-${testCase.rows}x${testCase.columns}.xlsx`);
      const generated = await generateSyntheticXlsx(path, { rows: testCase.rows, columns: testCase.columns });
      expect(generated.cells).toBe(testCase.rows * testCase.columns);
      const bytes = new Uint8Array(await readFile(path));

      const ws = await materializeXlsxWorksheet(bytes, 0, { xlsxWorksheetLimits: testCase.overrides });
      expect(ws.rows).toHaveLength(testCase.rows);
      const styles = indexStyles(bytes, ws);

      // Render first viewport, then the tail twice, observing the actual drawn text.
      const Base = skia.Canvas;
      vi.stubGlobal('OffscreenCanvas', class extends Base {});
      const canvas = new skia.Canvas(800, 400);
      const ctx = canvas.getContext('2d');
      const fillText = vi.spyOn(ctx, 'fillText');
      try {
        const target = canvas as unknown as OffscreenCanvas;
        const opts = { dpr: 1, width: 800, height: 400 };
        await renderWorksheetViewport({ ws, styles }, target, { row: 1, col: 1, rows: 20, cols: 5 }, opts);
        expect(fillText.mock.calls.map((call) => String(call[0]))).toContain('r1c1');
        for (let pass = 0; pass < 2; pass++) {
          fillText.mockClear();
          await renderWorksheetViewport({ ws, styles }, target, testCase.tailViewport, opts);
          expect(fillText.mock.calls.map((call) => String(call[0]))).toContain(testCase.tailText);
        }
        const data = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        let nonWhite = 0;
        for (let i = 0; i < data.length; i += 4) {
          if (data[i + 3] > 0 && (data[i] < 250 || data[i + 1] < 250 || data[i + 2] < 250)) nonWhite++;
        }
        expect(nonWhite).toBeGreaterThan(0);
      } finally {
        fillText.mockRestore();
        vi.unstubAllGlobals();
      }

      // Find over rendered display text.
      const find = new XlsxFindController(
        () => 1,
        () => 'Sheet1',
        async () => {
          const cells: FindCell[] = [];
          for (const row of ws.rows) {
            for (const cell of row.cells) {
              const text = formatCellValue(cell, styles, null, ws.date1904);
              if (text) cells.push({ row: row.index, col: cell.col, text });
            }
          }
          return cells;
        },
      );
      const matches = await find.find(testCase.tailText);
      expect(matches.length).toBeGreaterThanOrEqual(1);
      expect(matches[matches.length - 1].location.ref).toBe(testCase.tailRef);

      // Copy: tail selection writes the value; the full sheet stays independently too-large only when >250k.
      const writeText = vi.fn(async (_text: string) => {});
      const clipboard = { writeText } as unknown as Clipboard;
      let selection: XlsxSelectionState | null = selectionStateFromReference(testCase.tailRef);
      const copy = new CopyController({
        worksheet: () => ws,
        selection: () => selection,
        workbook: () => null,
        clipboard: () => clipboard,
      });
      await expect(copy.copy()).resolves.toMatchObject({ status: 'copied', cellCount: 1 });
      expect(writeText).toHaveBeenCalledWith(testCase.tailText);

      if (generated.cells > 250_000) {
        writeText.mockClear();
        selection = selectionStateFromReference(`A1:${testCase.tailRef}`);
        await expect(copy.copy()).resolves.toEqual({ status: 'too-large', limit: 'cells' });
        expect(writeText).not.toHaveBeenCalled();
      }
    }, LARGE_TIMEOUT);
  }
});
