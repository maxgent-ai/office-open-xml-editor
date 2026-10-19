import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';
import { legacyBundleBoundary, wasmAssetUrl } from '../../vite.config';
import { docxCompatibilityMetadata } from '../../scripts/docx-compatibility-build.mjs';

const __dirname = dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  define: { __OOXML_MODEL_SOURCES__: 'true' },
  // Library assets (parser WASM, workers) must resolve from the importing
  // module's URL, never from the hosting page's origin root, as in the root
  // build. Without it Vite 8.3 emits root-absolute "/<asset>" references.
  base: './',
  root: __dirname,
  plugins: [docxCompatibilityMetadata(), wasmAssetUrl(), wasm(), legacyBundleBoundary()],
  resolve: {
    alias: {
      '@ooxml-test-three-d-renderer': resolve(__dirname, '../../src/three-d.ts'),
      '@ooxml-test-region-map-renderer': resolve(__dirname, '../../src/region-map.ts'),
      '@ooxml-test-math-renderer': resolve(__dirname, '../../src/math.ts'),
      '@ooxml-test-tiff-renderer': resolve(__dirname, '../../src/tiff.ts'),
      '@ooxml-test-chart-ex-renderer': resolve(__dirname, '../../src/chart-ex.ts'),
    },
  },
  build: {
    // Serve public/ (sample fixtures) from the dev server for VRT, but don't
    // copy it into the published dist/.
    copyPublicDir: false,
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      formats: ['es', 'cjs'],
      fileName: (format) => `index.${format === 'es' ? 'mjs' : 'cjs'}`,
    },
    // Vite 8 (Rolldown) emits the WASM glue's top-level await natively, as
    // the xlsx and pptx package builds do; no down-level transform plugin.
    target: 'esnext',
  },
  server: {
    port: 5179,
    fs: {
      allow: [__dirname, resolve(__dirname, '../..')],
    },
  },
  worker: {
    format: 'es',
    plugins: () => [docxCompatibilityMetadata(), wasmAssetUrl(), wasm(), legacyBundleBoundary()],
    rollupOptions: {
      output: { assetFileNames: '[name][extname]' },
    },
  },
});
