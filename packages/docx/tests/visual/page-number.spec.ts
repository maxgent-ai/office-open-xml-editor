import { test, expect } from '@playwright/test';
import { existsSync } from 'node:fs';

// ECMA-376 §17.6.12 on REAL private samples, against Word PDF ground truth
// (measured with pdftotext on each sample's Word PDF). Both documents carry
// `<w:pgNumType>` plus a footer PAGE field, and Word prints SEQUENTIAL footers:
//
// - sample-12: one section with `w:fmt="numberInDash"`; Word prints -1- … -8-
//   over 8 pages (the dashes are separate runs, so the page token is the digit).
// - sample-1: nextPage section `w:start="1"`, then a continuous section with
//   `w:start="2"`. The continuous section begins exactly AT a page boundary, so
//   its restart shows 2, which equals the natural continuation; Word prints
//   1 … 5 over 5 pages.
//
// The displayed footer number therefore equals the physical page number. A
// VISIBLE continuous restart (the #804 `w:start="50"` case, where the shared
// page counts as the section's first page) is pinned deterministically in
// page-number-field-render.test.ts.
//
// Skips gracefully when the (gitignored) sample is absent.
const CASES: { file: string; pageCount: number; width: number; expected: string[] }[] = [
  {
    file: 'private/docx/sample-12', pageCount: 8, width: 595,
    expected: ['1', '2', '3', '4', '5', '6', '7', '8'],
  },
  { file: 'private/docx/sample-1', pageCount: 5, width: 595, expected: ['1', '2', '3', '4', '5'] },
];

test.describe('page-number restart non-regression (§17.6.12)', () => {
  for (const { file, pageCount, width, expected } of CASES) {
    test(`${file}: footer PAGE numbers stay sequential`, async ({ page }) => {
      if (!existsSync(`${process.cwd()}/public/${file}.docx`)) {
        test.skip(true, `gitignored sample ${file}.docx not present`);
      }
      // Load a page in the Vite module graph first so a dynamic `import('/src/...')`
      // inside page.evaluate resolves against the dev server (the fixture imports it).
      await page.goto(`/tests/visual/fixture.html?file=demo/sample-1.docx&page=0&width=595`);
      await page.waitForFunction(
        () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
        { timeout: 30_000 },
      );
      // Render each page and collect the bottom-most SHORT text run (the footer
      // page number). A footer number is a 1–3 char numeric/roman/letter token near
      // the page bottom, so we take the lowest run whose text is <= 4 chars.
      const numbers = await page.evaluate(
        async ({ file, pageCount, width }) => {
          const { DocxDocument } = await import('/src/index.ts');
          const doc = await DocxDocument.load('/' + file + '.docx');
          const out: (string | null)[] = [];
          for (let i = 0; i < pageCount; i++) {
            const canvas = document.createElement('canvas');
            const runs: { text: string; y: number }[] = [];
            await doc.renderPage(canvas, i, {
              width, dpr: 1,
              onTextRun: (r: { text: string; y: number }) => runs.push({ text: r.text, y: r.y }),
            });
            // Footer page number = the lowest-on-page run whose trimmed text is a
            // pure page-number token (Arabic digits, or roman/letter glyphs). This
            // excludes footnote/dagger marks (†) and body text. Prefer the bottom-most.
            const isPageToken = (t: string) => /^[0-9]{1,4}$/.test(t) || /^[ivxlcdmIVXLCDM]{1,4}$/.test(t) || /^[a-zA-Z]{1,3}$/.test(t);
            const tokens = runs
              .map((r) => ({ text: r.text.trim(), y: r.y }))
              .filter((r) => isPageToken(r.text))
              .sort((a, b) => b.y - a.y);
            out.push(tokens.length ? tokens[0].text : null);
          }
          return out;
        },
        { file, pageCount, width },
      );
      // The first `expected.length` pages carry a footer; assert those.
      for (let i = 0; i < expected.length; i++) {
        expect(numbers[i], `${file} page ${i + 1} footer number`).toBe(expected[i]);
      }
    });
  }
});
