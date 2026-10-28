import { describe, expect, it } from 'vitest';
import {
  PptxFontPreloadAccumulator,
  pptxSlideCjkFallback,
  pptxSlideOfficeFontRequests,
} from './font-requests';
import type { Presentation, Slide } from './types';

// DrawingML a:cs is an independently authored font slot (ECMA-376
// §21.1.2.3.1), already consumed by run segmentation. Loading must preserve
// its resolved face just as it preserves a:latin, a:ea and a:sym.
describe('PPTX complex-script slot loading', () => {
  const textBody = {
    defaultBold: true, defaultItalic: true,
    paragraphs: [{ bullet: { type: 'none' },
      runs: [{ type: 'text', text: 'שלום', lang: 'he-IL',
        fontFamily: 'Arial', fontFamilyCs: 'Calibri' }],
    }],
  };

  it.each(['shape', 'table'] as const)('retains authored CS faces in %s preload names', (kind) => {
    const slide = { elements: [kind === 'shape'
      ? { type: 'shape', textBody }
      : { type: 'table', rows: [{ cells: [{ textBody }] }] },
    ] } as unknown as Slide;
    const accumulator = new PptxFontPreloadAccumulator(null, null);
    accumulator.addSlide(slide);
    expect(accumulator.names()).toContain('Calibri');
    expect(pptxSlideOfficeFontRequests(slide, null, null))
      .toContainEqual({ family: 'Calibri', weight: 700, style: 'italic' });
  });
});

describe('PPTX exact Office face requests', () => {
  it('does not pin a theme font when no run or inherited style names a face', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: { paragraphs: [{
        defFontFamily: null, bullet: { type: 'none' },
        runs: [{ type: 'text', text: 'unstyled', fontFamily: null }],
      }] } }],
    } as unknown as Slide;
    expect(pptxSlideOfficeFontRequests(slide, 'Calibri Light', 'Calibri')).toEqual([]);
  });

  it('resolves theme, paragraph, and run style for every authored face', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: {
        defaultBold: false, defaultItalic: false,
        paragraphs: [{
          defFontFamily: '+mn-lt', defBold: false, defItalic: false,
          bullet: { type: 'none' },
          runs: [
            { type: 'text', text: 'regular', fontFamily: null, bold: null, italic: null },
            { type: 'text', text: 'bold', fontFamily: null, bold: true, italic: null },
            { type: 'text', text: 'other', fontFamily: 'Arial', bold: true, italic: true },
          ],
        }],
      } }],
    } as unknown as Slide;
    // Every authored face is probed (no family-name gate), so line metrics can
    // admit catalog geometry only for faces proven present.
    expect(pptxSlideOfficeFontRequests(slide, 'Cambria', 'Calibri')).toEqual([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', presenceOnly: true },
      { family: 'Calibri', weight: 700, style: 'normal' },
      { family: 'Arial', weight: 700, style: 'italic' },
      { family: 'Arial', presenceOnly: true },
    ]);
  });

  it('requests the face actually used by character and auto-number markers', () => {
    const slide = {
      elements: [{ type: 'shape', textBody: { paragraphs: [
        { defBold: true, defItalic: true, bullet: { type: 'char', char: '•', fontFamily: 'Calibri' }, runs: [] },
        { defBold: true, defItalic: true, bullet: { type: 'autoNum', numType: 'arabicPeriod', fontFamily: null },
          runs: [{ type: 'text', text: 'Item', fontFamily: 'Calibri', bold: true, italic: true }] },
      ] } }],
    } as unknown as Slide;
    expect(pptxSlideOfficeFontRequests(slide, null, 'Calibri')).toEqual(expect.arrayContaining([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 700, style: 'italic' },
    ]));
  });
});

describe('PPTX font request names', () => {
  it('includes slide-local paragraph and run families, not only the first theme fonts', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'shape',
        textBody: {
          paragraphs: [{
            defFontFamily: 'Franklin Gothic Medium',
            runs: [{
              type: 'text',
              text: 'Title',
              fontFamily: null,
              fontFamilyEa: 'Yu Gothic',
              fontFamilySym: null,
            }],
          }],
        },
      }],
    } as unknown as Slide;
    const accumulator = new PptxFontPreloadAccumulator('Aptos Display', 'Aptos');
    accumulator.addSlide(slide);

    expect(accumulator.names()).toEqual(expect.arrayContaining([
      'Aptos Display',
      'Aptos',
      'Franklin Gothic Medium',
      'Yu Gothic',
    ]));
  });
});

describe('PptxFontPreloadAccumulator', () => {
  it('preserves full-presentation shape, table, and chart text semantics incrementally', () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [
        { type: 'shape', textBody: { paragraphs: [{ runs: [{ type: 'text', text: '日本語' }] }] } },
        { type: 'table', rows: [{ cells: [{ textBody: { paragraphs: [{ runs: [{ type: 'text', text: 'العربية' }] }] } }] }] },
        {
          type: 'chart',
          chart: {
            title: 'Заголовок',
            categories: ['หมวด'],
            series: [{ name: 'סדרה' }],
          },
        },
      ],
    } as unknown as Slide;
    const pres: Presentation = {
      slideWidth: 1,
      slideHeight: 1,
      slides: [slide],
      defaultTextColor: null,
      majorFont: 'Yu Gothic',
      minorFont: 'Aptos',
    };
    const incremental = new PptxFontPreloadAccumulator(pres.majorFont, pres.minorFont);
    incremental.addSlide(slide);
    expect(incremental.names()).toEqual([
      'Yu Gothic', 'Aptos',
      'Noto Sans JP', 'Noto Serif JP',
      'Noto Sans', 'Noto Serif',
      'Noto Naskh Arabic', 'Noto Sans Arabic',
      'Noto Sans Thai',
      'Noto Sans Hebrew', 'Noto Serif Hebrew',
    ]);
  });
});


it('keeps the union of stable per-slide preferences during progressive preflight', () => {
  const slideWith = (text: string) => ({ elements: [{ type: 'shape', textBody: {
    paragraphs: [{ runs: [{ type: 'text', text }] }],
  } }] } as Slide);
  const han = slideWith('漢字');
  const kana = slideWith('漢字かな');
  const korean = slideWith('漢字한글');
  const fonts = new PptxFontPreloadAccumulator('Calibri', 'Calibri', undefined, undefined, 'sc');
  fonts.addSlide(han);
  expect(fonts.names()).toContain('Noto Sans SC');
  const next = fonts.withSlide(kana);
  expect(next.names()).toContain('Noto Sans SC');
  expect(next.names()).toContain('Noto Sans JP');
  expect(fonts.names()).not.toContain('Noto Sans JP');
  expect(pptxSlideCjkFallback(han, 'Calibri', 'Calibri', 'sc')).toBe('sc');
  expect(pptxSlideCjkFallback(kana, 'Calibri', 'Calibri', 'sc')).toBe('jp');
  expect(pptxSlideCjkFallback(korean, 'Calibri', 'Calibri', 'sc')).toBe('kr');
});

it.each(['shape', 'table'])('preloads override defaults and %s equation preview fonts in both render modes', (type) => {
  const slide = {
    theme: { majorFont: 'Override Heading', minorFont: 'Override Body', defaultTextColor: null, hlinkColor: null },
    elements: [{ type, textBody: null, rows: [], mathFallback: { before: 0, after: 0,
      elements: [{ type: 'shape', textBody: { paragraphs: [{ bullet: { type: 'none' },
        runs: [{ type: 'text', text: 'Preview', fontFamily: 'Calibri', bold: true }] }] } }],
    } }],
  } as unknown as Slide;
  const accumulator = new PptxFontPreloadAccumulator('Parent Heading', 'Parent Body');
  accumulator.addSlide(slide);
  expect(accumulator.names()).toEqual(expect.arrayContaining(['Override Heading', 'Override Body', 'Calibri']));
  expect(pptxSlideOfficeFontRequests(slide, 'Parent Heading', 'Parent Body'))
    .toEqual([{ family: 'Calibri', weight: 700, style: 'normal' }, { family: 'Calibri', presenceOnly: true }]);
});

it('preloads script fonts present only in an inline equation branch preview', () => {
  const slide = { elements: [{ type: 'shape', textBody: { paragraphs: [{ runs: [],
    mathFallbacks: [{ start: 0, end: 2, text: '漢字かな' }],
  }] } }] } as unknown as Slide;
  const fonts = new PptxFontPreloadAccumulator('Calibri', 'Calibri');
  fonts.addSlide(slide);
  expect(fonts.names()).toContain('Noto Sans JP');
});
