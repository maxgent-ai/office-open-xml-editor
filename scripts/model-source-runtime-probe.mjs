import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import {
  materializeDocxDocument,
  materializeXlsxWorkbookIndex,
  materializePptxPresentation,
  openDocxDocument,
  openXlsxWorkbook,
  openPptxPresentation,
} from '../dist/node.mjs';

const sample = (format) => readFileSync(new URL(
  `../packages/${format}/public/demo/sample-1.${format}`, import.meta.url,
));

const requireNodePackage = createRequire(new URL('../packages/node/package.json', import.meta.url));
const { Canvas, loadImage } = requireNodePackage('skia-canvas');
const factory = {
  createCanvas: (width, height) => new Canvas(width, height),
  loadImage,
};
const docx = sample('docx');
const xlsx = sample('xlsx');
const pptx = sample('pptx');
await materializeDocxDocument(docx);
await materializeXlsxWorkbookIndex(xlsx);
await materializePptxPresentation(pptx);
for (const open of [
  () => openDocxDocument(docx, { factory }),
  () => openXlsxWorkbook(xlsx),
  () => openPptxPresentation(pptx),
]) {
  const session = await open();
  await session.close();
}
// Let any unawaited import posted during the ordinary load reach the loader.
await new Promise((resolve) => setTimeout(resolve, 50));
