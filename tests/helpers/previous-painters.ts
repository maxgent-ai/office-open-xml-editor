import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { createRequire } from 'node:module';
import { join, resolve } from 'node:path';
import { build } from 'rolldown';

// The same immutable main revision used by the original pixel digests. Render
// it with the current host Canvas instead: Skia 2.0.2's native pixels differ
// between macOS and Linux, even for text-free patterns and curves. CI checks
// out full history; a missing revision fails closed rather than skipping.
const revision = '71c3e209692b73173b5a8d8037182fb2e76b4d07';
export interface PreviousPainters {
  paintDrawingMLShape: typeof import('@silurus/ooxml-core')['paintDrawingMLShape'];
  paintDrawingLayout: typeof import('../../packages/docx/src/paint/canvas-drawing')['paintDrawingLayout'];
  renderSlide: typeof import('../../packages/pptx/src/renderer')['renderSlide'];
  renderViewport: typeof import('../../packages/xlsx/src/renderer')['renderViewport'];
}

let painters: Promise<PreviousPainters> | undefined;
export function loadPreviousPainters(): Promise<PreviousPainters> {
  return painters ??= buildPreviousPainters();
}

async function buildPreviousPainters(): Promise<PreviousPainters> {
  const root = resolve(import.meta.dirname, '../..');
  const scratch = await mkdtemp(join(tmpdir(), 'ooxml-previous-painters-'));
  try {
    const packages = ['core', 'docx', 'pptx', 'xlsx'];
    const archive = execFileSync('git', ['archive', revision,
      ...packages.flatMap(name => [`packages/${name}/src`, `packages/${name}/package.json`])],
    { cwd: root, maxBuffer: 64 * 1024 * 1024 });
    execFileSync('tar', ['-x', '-C', scratch], { input: archive });
    await symlink(join(root, 'node_modules'), join(scratch, 'node_modules'));
    const core = JSON.parse(await readFile(join(scratch, 'packages/core/package.json'), 'utf8'));
    const alias = Object.fromEntries(Object.entries(core.exports as Record<string, string>)
      .filter(([key]) => !key.includes('*'))
      .map(([key, target]) => ['@silurus/ooxml-core' + (key === '.' ? '' : key.slice(1)),
        resolve(scratch, 'packages/core', target)]));
    const entry = join(scratch, 'entry.ts');
    await writeFile(entry, `
      export { paintDrawingMLShape } from './packages/core/src/shape/drawingml-shape';
      export { paintDrawingLayout } from './packages/docx/src/paint/canvas-drawing';
      export { renderSlide } from './packages/pptx/src/renderer';
      export { renderViewport } from './packages/xlsx/src/renderer';
    `);
    const output = join(scratch, 'painters.cjs');
    await build({ input: entry, resolve: { alias }, platform: 'node',
      transform: { define: { __OOXML_MODEL_SOURCES__: 'true' } },
      output: { file: output, format: 'cjs', codeSplitting: false },
    });
    // Native require keeps Vitest from resolving baseline imports back into
    // the candidate source graph. The bundle includes all painter dependencies.
    return createRequire(import.meta.url)(output) as PreviousPainters;
  } finally {
    // rm does not follow the dependency symlink; only this owned temporary
    // archive/bundle is removed, on success and failure alike.
    await rm(scratch, { recursive: true, force: true });
  }
}
