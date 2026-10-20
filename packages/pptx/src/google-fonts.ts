import type { CjkLang } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkFallbackForText,
  scriptPreloadNamesForText,
  GOOGLE_FONT_SUBSTITUTES,
  SCRIPT_GOOGLE_FONTS,
  type FontPreloadEntry,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { Presentation, Slide, SlideElement, TextBody } from './types';
import { officeFaceWeight } from './powerpoint-line-metrics.js';

/** Theme-referenced typefaces commonly used by PPTX templates. Keys are
 *  lower-cased family names.
 *
 *  {@link GOOGLE_FONT_SUBSTITUTES} supplies base text-face substitutes
 *  (Calibri → Carlito, Cambria → Caladea), the popular free
 *  web fonts and the Arabic Noto fallbacks — shared with docx/xlsx; the
 *  renderer puts each substitute into the canvas font stack so a missing Office
 *  font degrades to an advance-width alternative instead of a generic system serif/sans.
 *  {@link SCRIPT_GOOGLE_FONTS} adds the CJK / Cyrillic / Thai / Devanagari /
 *  Hebrew Noto faces (CJK ordered by document language). Both load only when
 *  `useGoogleFonts` is on. The exact local Calibri route is applied only to
 *  named font choices; the CSS alias for an unavailable face is likewise
 *  enabled only by that explicit opt-in. PPTX has no format-specific additions. */
export const PPTX_GOOGLE_FONTS: Record<string, FontPreloadEntry> = {
  ...GOOGLE_FONT_SUBSTITUTES,
  ...SCRIPT_GOOGLE_FONTS,
};

/** Include authored equation previews in font preparation, so runtime recovery
 * uses the same fonts in main-thread and worker rendering. Iterators retain
 * only the branch stack rather than flattening a second slide model. */
function* fontElements(elements: readonly SlideElement[]): Generator<SlideElement> {
  const pending: Iterator<SlideElement>[] = [elements[Symbol.iterator]()];
  while (pending.length) {
    const next = pending[pending.length - 1].next();
    if (next.done) { pending.pop(); continue; }
    const element = next.value;
    yield element;
    if ((element.type === 'shape' || element.type === 'table') && element.mathFallback) {
      pending.push(element.mathFallback.elements[Symbol.iterator]());
    }
  }
}

/** Yield every painted text string in a text body (paragraph runs). */
function* textBodyRuns(body: TextBody | null | undefined): Generator<string> {
  for (const p of body?.paragraphs ?? []) {
    for (const fallback of p.mathFallbacks ?? []) yield fallback.text;
    for (const r of p.runs) {
      if (r.type === 'text') yield r.text;
      if (r.type === 'math' && r.fallbackText) yield r.fallbackText;
    }
  }
}

/** Yield every explicitly resolved family carried by one rendered text body. */
function* textBodyFontFamilies(body: TextBody | null | undefined): Generator<string> {
  for (const paragraph of body?.paragraphs ?? []) {
    if (paragraph.defFontFamily) yield paragraph.defFontFamily;
    for (const run of paragraph.runs) {
      if (run.type !== 'text') continue;
      if (run.fontFamily) yield run.fontFamily;
      if (run.fontFamilyEa) yield run.fontFamilyEa;
      // ECMA-376 §21.1.2.3.1: a:cs names an independently resolved run face.
      // The existing renderer consumes it for complex-script segments;
      // preload must retain that independently authored face.
      if (run.fontFamilyCs) yield run.fontFamilyCs;
      if (run.fontFamilySym) yield run.fontFamilySym;
    }
  }
}

/** Yield every rendered text string in the presentation: shape text, table
 *  cell text and chart labels across all slides. Speaker notes and comments are
 *  not painted on the slide, so they are excluded (the renderer ignores them). */
export function* pptxSlideTextRuns(slide: Slide): Generator<string> {
  for (const el of fontElements(slide.elements)) {
    if (el.type === 'shape') {
      yield* textBodyRuns(el.textBody);
    } else if (el.type === 'table') {
      for (const row of el.rows) {
        for (const cell of row.cells) yield* textBodyRuns(cell.textBody);
      }
    } else if (el.type === 'chart') {
      if (el.chart.title) yield el.chart.title;
      for (const c of el.chart.categories) yield c;
      for (const s of el.chart.series) if (s.name) yield s.name;
    }
  }
}

/** Incremental PPTX adapter over core's single script-classification source. */
export class PptxFontPreloadAccumulator {
  private readonly scripts: ScriptPreloadAccumulator;
  private readonly families: Set<string>;

  constructor(
    private readonly majorFont: string | null,
    private readonly minorFont: string | null,
    scripts?: ScriptPreloadAccumulator,
    families?: Set<string>,
    private readonly fallback?: CjkLang,
    private readonly scriptNames = new Set<string>(),
  ) {
    const cjkLang = classifyCjkFont(majorFont) ?? classifyCjkFont(minorFont) ?? fallback ?? null;
    this.scripts = scripts ?? new ScriptPreloadAccumulator(cjkLang);
    this.families = families ?? new Set();
    if (majorFont) this.families.add(majorFont);
    if (minorFont) this.families.add(minorFont);
  }

  addSlide(slide: Slide): void {
    if (slide.theme?.majorFont) this.families.add(slide.theme.majorFont);
    if (slide.theme?.minorFont) this.families.add(slide.theme.minorFont);
    this.scripts.addText(pptxSlideTextRuns(slide));
    // Keep the union of per-slide choices: later kana must not remove a Han-only
    // slide's SC preload after that slide has already been published.
    for (const name of scriptPreloadNamesForText(pptxSlideTextRuns(slide),
      classifyCjkFont(slide.theme ? slide.theme.majorFont : this.majorFont)
        ?? classifyCjkFont(slide.theme ? slide.theme.minorFont : this.minorFont) ?? this.fallback ?? null)) {
      this.scriptNames.add(name);
    }
    for (const el of fontElements(slide.elements)) {
      if (el.type === 'shape') {
        for (const family of textBodyFontFamilies(el.textBody)) this.families.add(family);
      } else if (el.type === 'table') {
        for (const row of el.rows) {
          for (const cell of row.cells) {
            for (const family of textBodyFontFamilies(cell.textBody)) this.families.add(family);
          }
        }
      }
    }
  }

  names(): (string | null)[] {
    return [...new Set([...this.families, ...this.scripts.names(), ...this.scriptNames])];
  }

  withSlide(slide: Slide): PptxFontPreloadAccumulator {
    const candidate = new PptxFontPreloadAccumulator(
      this.majorFont,
      this.minorFont,
      this.scripts.clone(),
      new Set(this.families),
      this.fallback,
      new Set(this.scriptNames),
    );
    candidate.addSlide(slide);
    return candidate;
  }
}

const CSS_GENERIC_FAMILY_NAMES = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

/** Authored face tuples actually painted by this slide. DrawingML run properties
 * inherit through paragraph/body defaults (§21.1.2.3.9); theme tokens resolve
 * against the same major/minor values used by the renderer. Loading is delayed
 * until the slide is requested, so unused slides and styles fetch no assets. */
export function pptxSlideOfficeFontRequests(
  slide: Slide,
  majorFont: string | null,
  minorFont: string | null,
): OfficeFontFallbackRequest[] {
  if (slide.theme) { majorFont = slide.theme.majorFont; minorFont = slide.theme.minorFont; }
  const requests = new Map<string, OfficeFontFallbackRequest>();
  const resolved = (family: string | null | undefined): string | null => {
    // A missing family has no authored or inherited font resource. The theme
    // minor face is the renderer's CSS fallback, but it must not be treated as
    // a document request for an exact Office face. Explicit +mn-* still is.
    if (!family) return null;
    if (family.startsWith('+mn-')) return minorFont;
    if (family.startsWith('+mj-')) return majorFont;
    return family.split(',')[0]?.trim() ?? null;
  };
  const add = (family: string | null | undefined, bold: boolean, italic: boolean) => {
    // Every authored face is probed, so layout can admit catalog geometry
    // only for faces proven present (no family-name gate).
    const name = resolved(family);
    if (!name || CSS_GENERIC_FAMILY_NAMES.has(name.toLowerCase())) return;
    const weight = officeFaceWeight(name, bold, italic);
    const style = italic ? 'italic' : 'normal';
    requests.set(`${name.toLowerCase()}:${weight}:${style}`, { family: name, weight, style });
    // Family presence is probed independently of the styled tuples a slide
    // uses, so an installed regular-only family stays present for bold runs.
    requests.set(`presence:${name.toLowerCase()}`, { family: name, presenceOnly: true });
  };
  const body = (textBody: TextBody | null | undefined) => {
    for (const paragraph of textBody?.paragraphs ?? []) {
      for (const run of paragraph.runs) {
        if (run.type === 'break') {
          // a:br carries its own character properties (CT_TextLineBreak); the
          // renderer sizes the break's line with that face.
          add(run.fontFamily ?? paragraph.defFontFamily,
            run.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false,
            run.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false);
          continue;
        }
        if (run.type !== 'text') continue;
        const bold = run.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false;
        const italic = run.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false;
        add(run.fontFamily ?? paragraph.defFontFamily, bold, italic);
        if (run.fontFamilyEa) add(run.fontFamilyEa, bold, italic);
        if (run.fontFamilyCs) add(run.fontFamilyCs, bold, italic);
        if (run.fontFamilySym) add(run.fontFamilySym, bold, italic);
      }
      // a:endParaRPr: the end mark contributes its face to the last line.
      const end = paragraph.endRunProperties;
      if (end) {
        add(end.fontFamily ?? paragraph.defFontFamily,
          end.bold ?? paragraph.defBold ?? textBody?.defaultBold ?? false,
          end.italic ?? paragraph.defItalic ?? textBody?.defaultItalic ?? false);
      }
      if (paragraph.bullet.type === 'char' || paragraph.bullet.type === 'autoNum') {
        // The marker renderer uses normal weight/style for both bullet kinds.
        // autoNum inherits the first text run's family when buFont is absent.
        const firstRun = paragraph.runs.find((run) => run.type === 'text' && !!run.fontFamily);
        const firstRunFamily = firstRun?.type === 'text' ? firstRun.fontFamily : null;
        add(paragraph.bullet.fontFamily ?? (paragraph.bullet.type === 'autoNum'
          ? firstRunFamily ?? paragraph.defFontFamily
          : undefined), false, false);
      }
    }
  };
  for (const element of fontElements(slide.elements)) {
    if (element.type === 'shape') body(element.textBody);
    if (element.type === 'table') {
      for (const row of element.rows) for (const cell of row.cells) body(cell.textBody);
    }
  }
  return [...requests.values()];
}


export function pptxSlideCjkFallback(slide: Slide, major: string | null, minor: string | null, fallback: CjkLang): CjkLang {
  return cjkFallbackForText(pptxSlideTextRuns(slide),
    classifyCjkFont(major) ?? classifyCjkFont(minor) ?? fallback);
}
