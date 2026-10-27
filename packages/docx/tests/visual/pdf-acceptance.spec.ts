import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { basename, resolve } from 'node:path';
import { expect, test } from '@playwright/test';
import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';

interface PageExpectation {
  readonly includes?: readonly string[];
  readonly excludes?: readonly string[];
  readonly minMatchPct?: number;
}

interface AcceptanceManifest {
  readonly pageCount?: number;
  readonly minMatchPct?: number;
  readonly pages?: Readonly<Record<string, PageExpectation>>;
}

interface BrowserRender {
  readonly pageCount: number;
  readonly pages: readonly Readonly<{
    dataUrl: string;
    text: string;
  }>[];
}

const enabled = process.env.LOCAL_PDF_ACCEPTANCE === '1';

function requiredEnvironment(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`${name} is required for the local Word-PDF acceptance gate`);
  return value;
}

function readManifest(path: string | undefined): AcceptanceManifest {
  if (!path) return {};
  return JSON.parse(readFileSync(resolve(path), 'utf8')) as AcceptanceManifest;
}

function normalizedText(value: string): string {
  return value.normalize('NFKC').replace(/\s+/gu, ' ').trim();
}

function paddedPng(source: PNG, width: number, height: number): PNG {
  if (source.width === width && source.height === height) return source;
  const result = new PNG({ width, height });
  result.data.fill(255);
  PNG.bitblt(
    source,
    result,
    0,
    0,
    Math.min(source.width, width),
    Math.min(source.height, height),
    0,
    0,
  );
  return result;
}

function pdfPagePngs(pdfPath: string, outputPrefix: string): readonly string[] {
  const probe = spawnSync('pdftoppm', ['-v'], { encoding: 'utf8' });
  if (probe.error || probe.status !== 0) {
    throw new Error('pdftoppm was not found. Install Poppler with: brew install poppler');
  }
  execFileSync('pdftoppm', ['-png', '-r', '72', pdfPath, outputPrefix], {
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  const directory = resolve(outputPrefix, '..');
  const prefix = basename(outputPrefix);
  return readdirSync(directory)
    .filter((name) => name.startsWith(`${prefix}-`) && name.endsWith('.png'))
    .sort((left, right) => {
      const pageNumber = (name: string) => Number(/-(\d+)\.png$/u.exec(name)?.[1] ?? 0);
      return pageNumber(left) - pageNumber(right);
    })
    .map((name) => resolve(directory, name));
}

test.describe('local Word PDF acceptance', () => {
  test.skip(!enabled, 'Set LOCAL_PDF_ACCEPTANCE=1 or run pnpm vrt:pdf explicitly');

  test('matches page allocation, semantic expectations, and rendered pixels', async ({ page }, testInfo) => {
    const pdfPath = resolve(requiredEnvironment('DOCX_WORD_PDF'));
    if (!existsSync(pdfPath)) throw new Error(`DOCX_WORD_PDF does not exist: ${pdfPath}`);
    const documentPath = requiredEnvironment('DOCX_WORD_FILE').replace(/^\/+/, '');
    const manifest = readManifest(process.env.DOCX_WORD_EXPECTATIONS);
    const configuredMinimum = Number(process.env.DOCX_WORD_MIN_MATCH ?? manifest.minMatchPct ?? 96);
    if (!Number.isFinite(configuredMinimum) || configuredMinimum < 0 || configuredMinimum > 100) {
      throw new Error(`Invalid minimum pixel match percentage: ${configuredMinimum}`);
    }

    const pdfPrefix = testInfo.outputPath('word-page');
    const referencePaths = pdfPagePngs(pdfPath, pdfPrefix);
    if (referencePaths.length === 0) throw new Error(`No pages were rendered from ${pdfPath}`);
    const referencePages = referencePaths.map((path) => PNG.sync.read(readFileSync(path)));
    const requestedWidth = referencePages[0]!.width;

    await page.goto('/tests/visual/fixture.html');
    const rendered = await page.evaluate(async ({ documentPath: path, width }) => {
      const { DocxDocument } = await import('/src/document.ts');
      const { math } = await import('/tests/visual/math-engine.ts');
      const document = await DocxDocument.load(`/${path}`, {
        math,
      });
      const pages = [];
      for (let pageIndex = 0; pageIndex < document.pageCount; pageIndex += 1) {
        const canvas = window.document.createElement('canvas');
        await document.renderPage(canvas, pageIndex, { width, dpr: 1 });
        const runs = await document.collectPageRuns(pageIndex, { width, dpr: 1 });
        pages.push({
          dataUrl: canvas.toDataURL('image/png'),
          text: runs.map((run) => run.text).join(' '),
        });
      }
      return { pageCount: document.pageCount, pages };
    }, { documentPath, width: requestedWidth }) as BrowserRender;

    const expectedPageCount = manifest.pageCount ?? referencePages.length;
    expect(referencePages, 'Word PDF page count').toHaveLength(expectedPageCount);
    // Keep the page-count assertion loud, but compare every page that exists on
    // both sides first. A pagination regression must still leave useful image
    // evidence for the shared prefix instead of failing before page 1 is diffed.
    const comparablePageCount = Math.min(referencePages.length, rendered.pages.length);

    for (let pageIndex = 0; pageIndex < comparablePageCount; pageIndex += 1) {
      const pageNumber = pageIndex + 1;
      const expectation = manifest.pages?.[String(pageNumber)];
      const text = normalizedText(rendered.pages[pageIndex]!.text);
      for (const included of expectation?.includes ?? []) {
        expect(text, `page ${pageNumber} must include ${JSON.stringify(included)}`)
          .toContain(normalizedText(included));
      }
      for (const excluded of expectation?.excludes ?? []) {
        expect(text, `page ${pageNumber} must exclude ${JSON.stringify(excluded)}`)
          .not.toContain(normalizedText(excluded));
      }

      const reference = referencePages[pageIndex]!;
      const actualBuffer = Buffer.from(
        rendered.pages[pageIndex]!.dataUrl.replace(/^data:image\/png;base64,/u, ''),
        'base64',
      );
      const actual = PNG.sync.read(actualBuffer);
      const widthDifference = Math.abs(reference.width - actual.width);
      const heightDifference = Math.abs(reference.height - actual.height);
      expect(widthDifference, `page ${pageNumber} width rounding difference`).toBeLessThanOrEqual(1);
      expect(heightDifference, `page ${pageNumber} height rounding difference`).toBeLessThanOrEqual(1);
      const width = Math.max(reference.width, actual.width);
      const height = Math.max(reference.height, actual.height);
      const referencePadded = paddedPng(reference, width, height);
      const actualPadded = paddedPng(actual, width, height);
      const diff = new PNG({ width, height });
      const differentPixels = pixelmatch(
        referencePadded.data,
        actualPadded.data,
        diff.data,
        width,
        height,
        { threshold: 0.2, includeAA: false },
      );
      const totalPixels = width * height;
      const matchPct = 100 - differentPixels / totalPixels * 100;
      const minimum = expectation?.minMatchPct ?? configuredMinimum;
      const actualPath = testInfo.outputPath(`page-${pageNumber}-actual.png`);
      const referencePath = testInfo.outputPath(`page-${pageNumber}-word.png`);
      const diffPath = testInfo.outputPath(`page-${pageNumber}-diff.png`);
      writeFileSync(actualPath, PNG.sync.write(actualPadded));
      writeFileSync(referencePath, PNG.sync.write(referencePadded));
      writeFileSync(diffPath, PNG.sync.write(diff));
      await testInfo.attach(`page-${pageNumber}-actual`, { path: actualPath, contentType: 'image/png' });
      await testInfo.attach(`page-${pageNumber}-word`, { path: referencePath, contentType: 'image/png' });
      await testInfo.attach(`page-${pageNumber}-diff`, { path: diffPath, contentType: 'image/png' });
      console.log(
        `page ${pageNumber}: match=${matchPct.toFixed(3)}% `
        + `diff=${differentPixels.toLocaleString()}/${totalPixels.toLocaleString()} px `
        + `minimum=${minimum.toFixed(3)}%`,
      );
      expect(matchPct, `page ${pageNumber} Word-PDF pixel match`).toBeGreaterThanOrEqual(minimum);
    }

    expect(rendered.pages, 'DOCX rendered page count').toHaveLength(expectedPageCount);
    expect(rendered.pageCount).toBe(expectedPageCount);
  });
});
