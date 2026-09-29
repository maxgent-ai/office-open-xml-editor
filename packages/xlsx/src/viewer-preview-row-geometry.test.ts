/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from './wasm/xlsx_parser.js';
import { XlsxViewer } from './viewer.js';
import { prepareXlsxViewerRowHeights } from './workbook.js';
import { WorksheetPreview } from './internal/worksheet-preview.js';
import { getGridGeometryForWorksheet } from './renderer.js';
import type { Row, ViewportRange, Worksheet } from './types.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

interface CursorArchive {
  open_sheet_cursor(index: number, name: string): void;
  pull_sheet_cursor(rows: number): Uint8Array;
  acknowledge_sheet_cursor_terminal(): void;
  free(): void;
}

// The generated bindings are git-ignored build output; CI builds them before
// `pnpm test`. Report the suite as skipped rather than silently passing when a
// local checkout has not built them.
const wasmReady = (() => {
  try {
    const wasm = xlsxWasm as unknown as { initSync(init: { module: BufferSource }): unknown };
    wasm.initSync({ module: readFileSync(new URL('./wasm/xlsx_parser_bg.wasm', import.meta.url)) });
    return true;
  } catch {
    return false;
  }
})();

const CRC_TABLE = Uint32Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** Minimal stored (uncompressed) ZIP of UTF-8 entries. */
function storedZip(files: ReadonlyArray<readonly [string, string]>): Uint8Array {
  const encoder = new TextEncoder();
  const local: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const header = (size: number, fill: (view: DataView) => void, tail: Uint8Array[]) => {
    const head = new Uint8Array(size);
    fill(new DataView(head.buffer));
    return [head, ...tail];
  };
  for (const [name, content] of files) {
    const nameBytes = encoder.encode(name);
    const data = encoder.encode(content);
    const crc = crc32(data);
    const entry = header(30, (v) => {
      v.setUint32(0, 0x04034b50, true); v.setUint16(4, 20, true);
      v.setUint32(14, crc, true); v.setUint32(18, data.length, true);
      v.setUint32(22, data.length, true); v.setUint16(26, nameBytes.length, true);
    }, [nameBytes, data]);
    central.push(...header(46, (v) => {
      v.setUint32(0, 0x02014b50, true); v.setUint16(4, 20, true); v.setUint16(6, 20, true);
      v.setUint32(16, crc, true); v.setUint32(20, data.length, true);
      v.setUint32(24, data.length, true); v.setUint16(28, nameBytes.length, true);
      v.setUint32(42, offset, true);
    }, [nameBytes]));
    local.push(...entry);
    offset += entry.reduce((total, part) => total + part.length, 0);
  }
  const centralSize = central.reduce((total, part) => total + part.length, 0);
  const end = header(22, (v) => {
    v.setUint32(0, 0x06054b50, true); v.setUint16(8, files.length, true);
    v.setUint16(10, files.length, true); v.setUint32(12, centralSize, true);
    v.setUint32(16, offset, true);
  }, []);
  const parts = [...local, ...central, ...end];
  const out = new Uint8Array(parts.reduce((total, part) => total + part.length, 0));
  let cursor = 0;
  for (const part of parts) { out.set(part, cursor); cursor += part.length; }
  return out;
}

function workbookBytes(sheetXml: string): Uint8Array {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '</Types>'],
    ['_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      + '</Relationships>'],
    ['xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'],
    ['xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
      + '</Relationships>'],
    ['xl/worksheets/sheet1.xml', sheetXml],
  ]);
}

type Unit =
  | { kind: 'preview'; worksheet: Worksheet | null; reason: null; maxRow: number; maxCol: number }
  | { kind: 'rows'; rows: Row[] }
  | { kind: 'finished'; worksheet: Worksheet };

/** The parser units for one worksheet, in pull order. */
function pullUnits(sheetXml: string): Unit[] {
  const Archive = (xlsxWasm as unknown as {
    XlsxArchive: new (data: Uint8Array, a: null, b: null, c: null) => CursorArchive;
  }).XlsxArchive;
  const archive = new Archive(workbookBytes(sheetXml), null, null, null);
  try {
    archive.open_sheet_cursor(0, 'Sheet1');
    const units: Unit[] = [];
    for (;;) {
      const unit = JSON.parse(new TextDecoder().decode(archive.pull_sheet_cursor(128))) as Unit;
      units.push(unit);
      if (unit.kind === 'finished') break;
    }
    archive.acknowledge_sheet_cursor_terminal();
    return units;
  } finally {
    archive.free();
  }
}

/** Row bands the viewer actually painted: the viewport it chose and the
 * offset and size of every row in it, from the viewer's own grid geometry. */
function frame(worksheet: Worksheet, range: ViewportRange): Array<[number, number, number]> {
  const axis = getGridGeometryForWorksheet(worksheet).row;
  return Array.from({ length: range.rows }, (_, i) => {
    const row = range.row + i;
    return [row, axis.offsetOf(row), axis.sizeOf(row)];
  });
}

/** Show a large worksheet through the viewer's provisional path and return
 * the first painted frame (from the first row chunk) and the frame painted
 * after the pull completes. */
async function firstAndCompletedFrames(sheetFormat: string, rows: string) {
  const units = pullUnits('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `${sheetFormat}<sheetData>${rows}</sheetData></worksheet>`);
  const [preview, ...rest] = units;
  if (preview.kind !== 'preview' || !preview.worksheet) throw new Error('expected an exact preview');
  const chunks = rest.filter((unit): unit is Extract<Unit, { kind: 'rows' }> => unit.kind === 'rows');
  const finished = rest.at(-1);
  if (finished?.kind !== 'finished') throw new Error('expected a terminal worksheet');
  expect(chunks.length).toBeGreaterThan(1);

  installDom();
  const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
  const engine = viewer as unknown as Record<string, unknown> & {
    wb: unknown;
    showSheet(index: number): Promise<void>;
    renderCurrentSheet(): Promise<void>;
    currentWorksheet: Worksheet | null;
    previewFallbackReason: unknown;
    canvasArea: { clientWidth: number; clientHeight: number };
  };
  for (const method of [
    'hideCommentPopup', 'hideValidationPanel', 'updateSelectionOverlay', 'updateTabActive',
    'buildCommentMap', 'buildHyperlinkMap', 'buildOutline', 'layoutGutters', 'updateSpacerSize',
    'resetHorizontalScroll', 'updateFindOverlay', 'emitViewportChange', 'scheduleRender',
    'scheduleSelectionContextNotification',
  ]) engine[method] = vi.fn();
  engine.canvasArea.clientWidth = 800;
  engine.canvasArea.clientHeight = 600;

  const progress = new WorksheetPreview([]);
  progress.preview(preview.worksheet, preview.reason, preview.maxRow, preview.maxCol);
  progress.append(chunks[0].rows);
  let complete!: (worksheet: Worksheet) => void;
  const completion = new Promise<Worksheet>((resolve) => { complete = resolve; });
  const frames: Array<Array<[number, number, number]>> = [];
  engine.wb = {
    sheetNames: ['Sheet1'],
    sheetCount: 1,
    destroy: vi.fn(),
    retainWorksheetReference: vi.fn(async () => vi.fn()),
    acquireWorksheetPreviewLease: async () => ({
      worksheet: progress.worksheet, release: vi.fn(), partial: true, completion,
      waitForRows: (row: number) => progress.waitFor(row),
    }),
    [prepareXlsxViewerRowHeights]: vi.fn(),
    renderViewport: async (_target: unknown, _index: number, range: ViewportRange) => {
      frames.push(frame(engine.currentWorksheet as Worksheet, range));
    },
  };

  try {
    await engine.showSheet(0);
    expect(engine.previewFallbackReason).toBeNull();
    expect(frames).toHaveLength(1);
    expect(frames[0].at(-1)?.[0]).toBeLessThanOrEqual(progress.coveredThrough);

    // Complete the pull exactly as the workbook does: the terminal worksheet
    // adopts the streamed row array.
    for (const chunk of chunks.slice(1)) progress.append(chunk.rows);
    const terminal = finished.worksheet;
    terminal.rows = progress.rows;
    progress.finish(terminal);
    complete(terminal);
    await completion;
    await Promise.resolve();
    await engine.renderCurrentSheet();
    expect(frames).toHaveLength(2);
    const [first, completed] = frames;
    const band = (row: number) => first.find(([index]) => index === row)?.[2];
    return { first, completed, band };
  } finally {
    viewer.destroy();
  }
}

function dataRows(from: number, to: number): string {
  let rows = '';
  for (let r = from; r <= to; r++) {
    rows += r % 3 === 0
      ? `<row r="${r}"><c r="A${r}"><v>${r}</v></c></row>`
      : `<row r="${r}" ht="18"><c r="A${r}"><v>${r}</v></c></row>`;
  }
  return rows;
}

describe.skipIf(!wasmReady)('XLSX provisional first frame row geometry', () => {
  it('keeps the default band of zeroHeight visible rows without @ht', async () => {
    // ECMA-376 §18.3.1.81: zeroHeight hides only unspecified rows. Row 2 is an
    // explicit visible row without @ht and keeps the default band; rows after
    // the last authored row stay hidden.
    const { first, completed, band } = await firstAndCompletedFrames(
      '<sheetFormatPr defaultRowHeight="22" zeroHeight="1"/>',
      '<row r="1" ht="30"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>'
        + dataRows(3, 2_000),
    );
    expect(first).toEqual(completed);
    expect(band(2)).toBeGreaterThan(0);
    expect(band(3)).toBeGreaterThan(0);
  });

  it('reads a lenient hidden spelling like the completed sheet', async () => {
    const { first, completed, band } = await firstAndCompletedFrames(
      '',
      '<row r="1"><c r="A1"><v>1</v></c></row><row r="2" hidden="True"><c r="A2"><v>2</v></c></row>'
        + dataRows(3, 2_000),
    );
    expect(first).toEqual(completed);
    expect(band(2)).toBe(0);
  });
});
