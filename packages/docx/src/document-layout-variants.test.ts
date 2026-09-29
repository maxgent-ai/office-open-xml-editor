import { describe, expect, it, vi } from 'vitest';
import type { DocxDocumentModel } from './types.js';
import {
  attachDocumentLayoutVariants,
  selectDocumentLayoutPage,
} from './layout/document-layout-variants.js';
import { layoutVariantStoreOf } from './layout/runtime-state.js';
import type { DocumentLayout, LayoutServices } from './layout/types.js';
import { ensureDocumentLayoutVariants } from './layout/document.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';

describe('document canonical layout variants', () => {
  it('keeps default metadata isolated and validates the selected keyed variant', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const model = {
      body: [], section: { pageWidth: 100, pageHeight: 200 },
    } as unknown as DocxDocumentModel;
    const layout = (count: number): DocumentLayout => ({
      pages: Array.from({ length: count }, (_, pageIndex) => ({ pageIndex }) as never),
      diagnostics: [],
    });
    const variants = attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services, defaultCurrentDateMs: 10,
      buildLayout: (options) => layout(options.currentDateMs === 10 ? 2 : 1),
    });

    expect(variants.store.defaultLayout.pages).toHaveLength(2);
    expect(() => variants.store.selectPage({ currentDateMs: 20 }, 1)).toThrow(/out of range/);
    expect(variants.store.defaultLayout.pages).toHaveLength(2);
    expect(layoutVariantStoreOf(services)).toBe(variants.store);
  });

  it('normalizes and selects one keyed page from the store attached to the actual services', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const model = {
      body: [], section: { pageWidth: 100, pageHeight: 200 },
    } as unknown as DocxDocumentModel;
    const layout = (date: number): DocumentLayout => ({
      pages: [{ pageIndex: date === 10 ? 0 : 1 } as never],
      diagnostics: [],
    });
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model), services, defaultCurrentDateMs: 10,
      buildLayout: (options) => layout(options.currentDateMs),
    });

    const selected = selectDocumentLayoutPage(services, {
      currentDate: new Date(20), defaultCurrentDateMs: 10,
    }, 0);

    expect(selected.options).toEqual({ currentDateMs: 20 });
    expect(selected.key).toBe(selectDocumentLayoutPage(services, {
      currentDate: 20, defaultCurrentDateMs: 10,
    }, 0).key);
    expect(selected.key).not.toBe(selectDocumentLayoutPage(services, {
      defaultCurrentDateMs: 10,
    }, 0).key);
    expect(selected.page).toBe(selected.layout.pages[0]);
    expect(() => selectDocumentLayoutPage({ ...services }, {
      currentDate: 20, defaultCurrentDateMs: 10,
    }, 0)).toThrow(/variant store.*not attached/i);
  });

  it('does not resolve the document body again after variants are attached', () => {
    const services = Object.freeze({
      text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' },
    }) as LayoutServices;
    const model = {
      body: [], section: { pageWidth: 100, pageHeight: 200 },
    } as unknown as DocxDocumentModel;
    attachDocumentLayoutVariants({
      source: layoutSourceStore(model),
      services,
      defaultCurrentDateMs: 10,
      buildLayout: () => ({ pages: [{ pageIndex: 0 } as never], diagnostics: [] }),
    });
    const resolveSource = vi.fn(() => {
      throw new Error('the retained variant must short-circuit document projection');
    });

    ensureDocumentLayoutVariants(services, 10, resolveSource);

    expect(resolveSource).not.toHaveBeenCalled();
  });
});
