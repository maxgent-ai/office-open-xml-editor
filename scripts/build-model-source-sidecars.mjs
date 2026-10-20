#!/usr/bin/env node
// Each optional worker-realm source owner is one self-contained ES module.
// A consumer bundler treats prebuilt workers and their URLs as opaque assets;
// sibling chunks would not be copied and cannot resolve from an inline worker
// blob. Keep these sidecars out of the ordinary OOXML graph and independent
// of relative JS imports.
import { resolve } from 'node:path';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { docxCompatibilityMetadata } from './docx-compatibility-build.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const owners = {
  docx: 'worker-document-source.ts',
  xlsx: 'worker-worksheet-source.ts',
  pptx: 'worker-presentation-source.ts',
};
// Source-built applications need these URL-loaded modules beside their emitted
// viewer chunks too. The site supplies its asset directory without requiring a
// full viewer-package build; package builds retain their existing destinations.
const args = process.argv.slice(2);
const outputOverride = args[0] === '--out-dir';
if (args.length > (outputOverride ? 2 : 1) || (outputOverride && !args[1])) {
  throw new Error('usage: build-model-source-sidecars.mjs [format | --out-dir directory]');
}
const selected = outputOverride ? undefined : args[0];
const outputDirectory = outputOverride ? resolve(args[1]) : undefined;
if (selected !== undefined && !(selected in owners)) {
  throw new Error(`unknown model-source sidecar format: ${selected}`);
}

for (const format of selected ? [selected] : Object.keys(owners)) {
  const packageBuild = selected !== undefined;
  await build({
    configFile: false,
    plugins: [docxCompatibilityMetadata()],
    base: './',
    build: {
      outDir: outputDirectory ?? (packageBuild ? resolve(root, 'packages', format, 'dist') : resolve(root, 'dist')),
      emptyOutDir: false,
      copyPublicDir: false,
      target: 'esnext',
      lib: {
        entry: packageBuild
          ? resolve(root, 'packages', format, 'src', 'internal', owners[format])
          : resolve(root, 'src', `${format}-source-worker.ts`),
        formats: ['es'],
        fileName: () => `${format}-source-worker.mjs`,
      },
      rolldownOptions: { output: { codeSplitting: false } },
    },
  });
}
