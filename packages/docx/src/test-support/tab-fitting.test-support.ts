import { CONFORMANCE_CASES } from '../conformance/cases.js';
import { generateConformanceParts, storeZip } from '../conformance/generate.js';

export function documentBytes(
  indent: string,
  rtl: boolean,
  xPt: number,
  tab?: { alignment: string; count: number; text: string; positional?: boolean; relativeTo?: 'margin' | 'indent'; noFloat?: boolean; fontSizePt?: number; automatic?: boolean; prefix?: string; stop?: number; cells?: readonly string[]; mirrorIndents?: boolean },
): Uint8Array {
  const seed = CONFORMANCE_CASES.find(({ axes }) => axes.story === 'body'
    && axes.container === 'paragraph' && axes.object === 'floating');
  if (!seed) throw new Error('Missing floating conformance seed');
  const parts = new Map(generateConformanceParts({
    ...seed,
    expected: { ...seed.expected, targetText: tab?.text ?? 'word '.repeat(30).trim() },
    axes: { ...seed.axes, direction: rtl ? 'rtl' : 'ltr', paragraph: 'single',
      styleSource: 'direct', fontSource: 'direct', spacing: 'exact', anchorReference: 'margin' },
  }));
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let xml = decoder.decode(parts.get('word/document.xml'));
  xml = xml.replace('</w:pPr>', `<w:ind ${indent}/></w:pPr>`)
    .replace('<wp:posOffset>914400</wp:posOffset>', `<wp:posOffset>${xPt * 12700}</wp:posOffset>`)
    .replaceAll('cx="457200"', 'cx="2540000"')
    .replaceAll('cy="274320"', 'cy="1270000"');
  if (tab) {
    const tabXml = tab.positional
      ? `<w:ptab w:alignment="${tab.alignment}" w:relativeTo="${tab.relativeTo ?? 'margin'}" w:leader="none"/>`
      : '<w:tab/>'.repeat(tab.count);
    xml = xml.replace('<w:t', `${tab.prefix ? `<w:t>${tab.prefix}</w:t>` : ""}${tabXml}<w:t`)
      .replace('</w:pPr>', `<w:tabs><w:tab w:val="${tab.alignment}" w:pos="${tab.stop ?? 5600}"/></w:tabs></w:pPr>`);
    // CT_PPrBase requires tabs before bidi/spacing/ind (§A.1).
    xml = xml.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (_match, properties: string) => {
      const tabs = properties.match(/<w:tabs>[\s\S]*?<\/w:tabs>/)?.[0] ?? '';
      return `<w:pPr>${tabs}${properties.replace(tabs, '')}</w:pPr>`;
    });
    // The float's anchor must precede the tab cell, rather than count as
    // trailing content when the aligned tab measures that cell.
    const drawingRun = xml.match(/<w:r>\s*<w:drawing>[\s\S]*?<\/w:drawing>\s*<\/w:r>/)?.[0];
    if (!drawingRun) throw new Error('Missing floating drawing run');
    xml = xml.replace(drawingRun, '').replace('</w:pPr>', `</w:pPr>${drawingRun}`);
  }
  if (tab?.cells) {
    xml = xml.replace(/(<w:tab\/>)+(?=<w:t)/, '')
      .replace(/<w:t\b[^>]*>[\s\S]*?<\/w:t>/, tab.cells.map((cell) => `<w:t xml:space="preserve">${cell}</w:t>`).join('<w:tab/>'));
  }
  if (tab?.mirrorIndents) xml = xml.replace(/<w:ind[^>]*\/>/, '$&<w:mirrorIndents/>');
  if (tab?.fontSizePt) xml = xml.replace(/<w:sz(?:Cs)? w:val="\d+"\/>/g, (tag) => tag.replace(/\d+/, String((tab.fontSizePt ?? 12) * 2)));
  if (tab?.automatic) xml = xml.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/g, '');
  if (tab?.noFloat) xml = xml.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '');
  parts.set('word/document.xml', encoder.encode(xml));
  return storeZip(parts);
}

export const measureContext = (advancePt = 5) => ({
  font: '', letterSpacing: '0px', fontKerning: 'normal',
  measureText: (text: string) => ({ width: [...text].length * advancePt,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
}) as CanvasRenderingContext2D;

export const tabCases = [
  { alignment: 'left', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
  { alignment: 'left', count: 2, ltrX: 352, ltrWidth: 300, rtlX: 240, rtlWidth: 300 },
  { alignment: 'right', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
  { alignment: 'right', count: 2, ltrX: 332, ltrWidth: 280, rtlX: 260, rtlWidth: 280 },
  { alignment: 'center', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
  { alignment: 'center', count: 2, ltrX: 342, ltrWidth: 290, rtlX: 250, rtlWidth: 290 },
  { alignment: 'decimal', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
  { alignment: 'decimal', count: 2, ltrX: 342, ltrWidth: 290, rtlX: 250, rtlWidth: 290 },
  { alignment: 'left', count: 1, relativeTo: 'margin', ltrX: 72, ltrWidth: 20, rtlX: 520, rtlWidth: 20 },
  { alignment: 'left', count: 1, relativeTo: 'indent', ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
  { alignment: 'center', count: 1, relativeTo: 'margin', ltrX: 296, ltrWidth: 244, rtlX: 296, rtlWidth: 244 },
  { alignment: 'center', count: 1, relativeTo: 'indent', ltrX: 314, ltrWidth: 262, rtlX: 278, rtlWidth: 262 },
  { alignment: 'right', count: 1, relativeTo: 'margin', ltrX: 520, ltrWidth: 468, rtlX: 72, rtlWidth: 468 },
  { alignment: 'right', count: 1, relativeTo: 'indent', ltrX: 520, ltrWidth: 468, rtlX: 72, rtlWidth: 468 },
] as const;
export const matrix = tabCases.flatMap((tab, tabIndex) => [false, true].flatMap((rtl) =>
  (['none', 'left', 'right'] as const).map((float) => ({
    ...tab, tabIndex, rtl, float, kind: 'relativeTo' in tab ? `positional-${tab.relativeTo}` : 'ordinary',
  }))));

export const followOnContents = [
  { script: 'Latin words', content: 'word '.repeat(40).trim() },
  { script: 'CJK', content: '漢'.repeat(100) },
  { script: 'Thai dictionary', content: 'ภาษาไทย'.repeat(20) },
  { script: 'unbreakable', content: 'a'.repeat(100) },
  { script: 'mixed', content: 'word 漢字ภาษาไทย abc-def '.repeat(20).trim() },
];
export const fittingMatrix = [...matrix, ...[false, true].flatMap((rtl) =>
  (['none', 'left', 'right'] as const).map((float) => ({
    alignment: 'left', count: 1, rtl, float, kind: 'automatic', tabIndex: 0,
  })))].flatMap((entry) => followOnContents.map((content) => ({ ...entry, ...content })));
