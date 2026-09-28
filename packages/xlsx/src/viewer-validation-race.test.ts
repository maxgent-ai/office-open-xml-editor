import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import type { XlsxWorkbook } from './workbook.js';
import type { Worksheet } from './types.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

interface ValidationTestSeam {
  wb: XlsxWorkbook;
  selectionController: { select(cell: { row: number; col: number }): void };
  currentSheet: number;
  currentWorksheet: Worksheet;
  hideValidationPanel(): void;
  validation: {
    open(cell: { row: number; col: number }, formula: string): Promise<void>;
    toggle(): void;
    panelKey: string | null;
    panel: { style: { display: string } };
  };
}

function pendingWorkbook() {
  const pending: Array<(value: { kind: 'values'; values: string[] }) => void> = [];
  const workbook = {
    resolveValidationList: vi.fn(() => new Promise((resolve) => { pending.push(resolve); })),
    destroy: vi.fn(),
  } as unknown as XlsxWorkbook;
  return { workbook, pending };
}

function mountPendingValidationViewer() {
  installDom();
  const { workbook, pending } = pendingWorkbook();
  const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
  const seam = viewer as unknown as ValidationTestSeam;
  seam.wb = workbook;
  seam.currentSheet = 0;
  seam.currentWorksheet = {
    name: 'Validation',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    freezeRows: 0,
    freezeCols: 0,
    mergeCells: [],
    conditionalFormats: [],
    images: [],
    charts: [],
    dataValidations: [
      { validationType: 'list', sqref: 'A1', formula1: 'A1:A2' },
      { validationType: 'list', sqref: 'B1', formula1: 'B1:B2' },
    ],
  } as Worksheet;
  return { viewer, seam, pending };
}

describe('XlsxViewer validation-list request lifecycle', () => {
  it('a second click on the same pending arrow cancels instead of reopening', async () => {
    const { viewer, seam, pending } = mountPendingValidationViewer();
    seam.selectionController.select({ row: 1, col: 1 });

    seam.validation.toggle();
    expect(seam.validation.panelKey).toBe('1:1');
    seam.validation.toggle();
    expect(seam.validation.panelKey).toBeNull();
    pending[0]?.({ kind: 'values', values: ['A'] });
    await Promise.resolve();
    await Promise.resolve();

    expect(seam.validation.panel.style.display).toBe('none');
    viewer.destroy();
  });

  it('only the latest cell request may install the panel', async () => {
    const { viewer, seam, pending } = mountPendingValidationViewer();
    seam.selectionController.select({ row: 1, col: 1 });
    seam.validation.toggle();
    seam.selectionController.select({ row: 1, col: 2 });
    seam.validation.toggle();

    pending[1]?.({ kind: 'values', values: ['B'] });
    await Promise.resolve();
    await Promise.resolve();
    pending[0]?.({ kind: 'values', values: ['A'] });
    await Promise.resolve();
    await Promise.resolve();

    expect(seam.validation.panelKey).toBe('1:2');
    expect(seam.validation.panel.style.display).not.toBe('none');
    viewer.destroy();
  });

  it('hide cancels a pending resolution', async () => {
    const { viewer, seam, pending } = mountPendingValidationViewer();
    seam.selectionController.select({ row: 1, col: 1 });
    seam.validation.toggle();
    seam.hideValidationPanel();
    pending[0]?.({ kind: 'values', values: ['A'] });
    await Promise.resolve();
    await Promise.resolve();

    expect(seam.validation.panel.style.display).toBe('none');
    viewer.destroy();
  });

  it('does not install a document listener when resolution completes after destroy', async () => {
    const doc = installDom();
    let resolveList: (value: { kind: 'values'; values: string[] }) => void = () => undefined;
    const workbook = {
      resolveValidationList: () => new Promise((resolve) => { resolveList = resolve; }),
      destroy: vi.fn(),
    } as unknown as XlsxWorkbook;
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const seam = viewer as unknown as ValidationTestSeam;
    seam.wb = workbook;
    seam.currentSheet = 0;
    seam.selectionController.select({ row: 1, col: 1 });
    seam.validation.panelKey = '1:1';

    const opening = seam.validation.open({ row: 1, col: 1 }, 'A1:A2');
    viewer.destroy();
    resolveList({ kind: 'values', values: ['A', 'B'] });
    await opening;

    expect(doc.listenerCount('pointerdown')).toBe(0);
  });
});
