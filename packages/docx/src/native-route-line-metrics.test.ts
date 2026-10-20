import { describe, expect, it } from 'vitest';
import { createLayoutServices, nativeRouteLineRequests } from './layout-runtime.js';
import { buildSegments, type LayoutTextSeg } from './line-layout.js';
import { nativeRouteMetricKey } from './layout/font-service.js';
import type { DocRun, DocxDocumentModel } from './types.js';

function run(text: string, fields: Partial<Record<string, unknown>> = {}): DocRun {
  return {
    type: 'text', text, fontFamily: 'Absent Serif 1843', fontSize: 11, bold: false, italic: false,
    underline: false, strikethrough: false, color: null, ...fields,
  } as DocRun;
}

describe('native route enumeration', () => {
  it('enumerates every native family list layout requests, per language, slot and tuple', () => {
    const runs = [
      run('Latin text'),
      run('Bold', { bold: true }),
      run('Italic bold', { bold: true, italic: true }),
      run('漢字 and Latin', { fontFamilyEastAsia: 'Absent Mincho 1843', langEastAsia: 'zh-CN' }),
      run('漢字', { fontFamilyEastAsia: 'Absent Mincho 1843', langEastAsia: 'ja-JP' }),
      run('かな', { fontFamilyEastAsia: 'Absent Gothic 1843' }),
    ];
    const empty = { default: null, first: null, even: null };
    const model = {
      section: { pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
        marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36 },
      body: [{ type: 'paragraph', runs, defaultFontFamily: 'Absent Serif 1843' }],
      headers: empty, footers: empty,
      fontFamilyClasses: { 'Absent Serif 1843': 'roman', 'Absent Mincho 1843': 'roman' },
    } as unknown as DocxDocumentModel;
    const options = { cjkFallback: 'kr' as const };
    const enumerated = new Set(nativeRouteLineRequests(model, model, options)
      .map(request => nativeRouteMetricKey(request.familyList, request.weight, request.style)));
    const layoutServices = createLayoutServices(model, options);
    const requested = new Set<string>();
    for (const authored of runs) {
      for (const segment of buildSegments([authored], { pageIndex: 0, totalPages: 1, layoutServices })) {
        const text = segment as LayoutTextSeg;
        if (text.fontSource !== 'native' || !text.fontRoute) continue;
        requested.add(nativeRouteMetricKey(text.fontRoute.familyList, text.bold ? 700 : 400,
          text.italic ? 'italic' : 'normal'));
      }
    }
    // Regional lists (zh-CN, ja-JP, the kr fallback) and three tuples occur.
    expect(requested.size).toBeGreaterThanOrEqual(5);
    for (const key of requested) expect(enumerated).toContain(key);
  });
});
