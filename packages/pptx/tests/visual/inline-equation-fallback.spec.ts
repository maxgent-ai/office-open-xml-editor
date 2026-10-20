import { test, expect } from '@playwright/test';
import { storedZip } from '../../../../tests/fixtures/chart-ex-packages.mjs';

for (const kind of ['two children', 'one paragraph', 'nested choice', 'nested fallback', 'ignored wrapper', 'process-content wrapper']) {
  test(`inline equation recovery replaces the complete Choice with ${kind}`, async ({ page }) => {
    const equation = (text: string) => `<m:oMath><m:r><m:t>${text}</m:t></m:r></m:oMath>`;
    const choice = kind === 'two children'
      ? `<a14:m>${equation('x')}</a14:m><a14:m>${equation('y')}</a14:m>`
      : `<a14:m><m:oMathPara>${equation('x')}${equation('y')}</m:oMathPara></a14:m>`;
    const preview = '<a:r><a:t>complete preview</a:t></a:r>';
    const discarded = '<a:r><a:t>discarded preview</a:t></a:r>';
    const nestedPreview = `<mc:AlternateContent>
      <mc:Choice Requires="u">${discarded}</mc:Choice>
      <mc:Choice Requires="${kind === 'nested fallback' ? 'a14 u' : 'a14'}">${preview}</mc:Choice>
      <mc:Choice Requires="${kind === 'nested fallback' ? 'u' : 'a14'}">${discarded}</mc:Choice>
      <mc:Fallback>${kind === 'nested fallback' ? preview : discarded}</mc:Fallback>
    </mc:AlternateContent>`;
    const fallback = kind === 'ignored wrapper'
      ? `${preview}<u:wrap>${nestedPreview}${discarded}</u:wrap>`
      : kind === 'process-content wrapper'
        ? `<u:wrap>${nestedPreview}</u:wrap><u:other>${discarded}</u:other>`
        : kind.startsWith('nested') ? nestedPreview : preview;
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const bytes = storedZip([
      ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
        <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`],
      ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`],
      ['ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst>
        <p:sldId id="256" r:id="rSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`],
      ['ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rSlide" Type="${r}/slide" Target="slides/slide1.xml"/></Relationships>`],
      ['ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
        xmlns:u="urn:unsupported" mc:Ignorable="u"
        ${kind === 'process-content wrapper' ? 'mc:ProcessContent="u:wrap"' : ''}
        xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><p:cSld><p:spTree>
        <p:sp><p:nvSpPr><p:cNvPr id="2" name="Equation"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>before</a:t></a:r><mc:AlternateContent>
        <mc:Choice Requires="a14">${choice}</mc:Choice>
        <mc:Fallback>${fallback}</mc:Fallback>
        </mc:AlternateContent><a:r><a:t>after</a:t></a:r></a:p></p:txBody>
        </p:sp></p:spTree></p:cSld></p:sld>`],
    ]);
    await page.route('**/equation-probe.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html>' }));
    await page.goto('/equation-probe.html');
    const results = await page.evaluate(async (data) => {
      const wasm = await import('/src/wasm/pptx_parser.js');
      await wasm.default({ module_or_path: fetch('/src/wasm/pptx_parser_bg.wasm') });
      const { renderSlide } = await import('/src/renderer.ts') as typeof import('../../src/renderer.js');
      const results = [];
      for (const failure of ['all', 'last']) {
        // Parse fresh nodes each time so successful equation caches cannot
        // conceal a conversion failure in the next production render.
        const deck = JSON.parse(new TextDecoder().decode(wasm.parse_pptx(new Uint8Array(data))));
        const slide = deck.slides[0];
        const original = JSON.stringify(slide);
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
        const painted: string[] = [];
        let images = 0;
        const fillText = ctx.fillText.bind(ctx);
        ctx.fillText = (text, x, y) => { painted.push(text); fillText(text, x, y); };
        const drawImage = ctx.drawImage.bind(ctx);
        ctx.drawImage = ((...args: Parameters<CanvasRenderingContext2D['drawImage']>) => {
          images++;
          drawImage(...args);
        }) as CanvasRenderingContext2D['drawImage'];
        await renderSlide(canvas, slide, deck.slideWidth, deck.slideHeight, { width: 960, dpr: 1,
          math: { loadMathJax: async () => {}, mathMLToSvg: async xml => {
            if (failure === 'all' || xml.includes('>y<')) throw new Error('conversion');
            return { svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10"/></svg>',
              widthEm: 1, ascentEm: 1, descentEm: 0 };
          } },
        });
        results.push({ failure, painted, images, unchanged: original === JSON.stringify(slide) });
      }
      return results;
    }, [...bytes]);
    for (const result of results) {
      expect(result, JSON.stringify(result)).toMatchObject({
        painted: ['before', 'complete preview', 'after'], images: 0, unchanged: true,
      });
    }
  });
}
