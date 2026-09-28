/**
 * The VS Code webview always renders on the main thread. Replace both worker
 * entry mechanisms before esbuild traverses them: the lazy host used by the
 * public viewers and the Vite `?worker&inline` parser worker imports retained
 * by their main-thread modules.
 *
 * Besides keeping unused workers out of the VSIX, this prevents esbuild from
 * entering worker-only modules that intentionally use Vite asset queries such
 * as `mathjax-stix2.js?url`.
 */
export const mainThreadOnlyWorkerStubs = {
  name: 'main-thread-only-worker-stubs',
  setup(build) {
    // The shared worker barrel is also imported by main-thread code for resource
    // limits and error transport. Replace only its renderer loader re-export;
    // otherwise esbuild follows the loader's dynamic imports even though the
    // exported functions are unused by this webview.
    build.onResolve({ filter: /^\.\/renderer-module\.js$/ }, (args) => {
      const importer = args.importer.replaceAll('\\', '/');
      if (!importer.endsWith('/packages/core/src/worker/index.ts')) return null;
      return { path: args.path, namespace: 'stub-worker-renderer-module' };
    });
    build.onLoad({ filter: /.*/, namespace: 'stub-worker-renderer-module' }, () => ({
      contents:
        "export async function loadWorkerRenderer() { throw new Error('[ooxml] worker renderer loading is not available in the VS Code extension (main-thread only)'); }" +
        " export async function loadWorkerRenderers() { return {}; }",
      loader: 'js',
    }));

    build.onResolve({ filter: /render-worker(?:-source)?-host(?:\.js)?$/ }, (args) => ({
      path: args.path,
      namespace: 'stub-render-worker-host',
    }));
    build.onLoad({ filter: /.*/, namespace: 'stub-render-worker-host' }, () => ({
      contents:
        "export function createRenderWorker() {" +
        " throw new Error('[ooxml] worker rendering is not available in the VS Code extension (main-thread only)'); }",
      loader: 'js',
    }));

    build.onResolve({ filter: /\?worker&inline$/ }, (args) => ({
      path: args.path,
      namespace: 'stub-inline-worker',
    }));
    build.onLoad({ filter: /.*/, namespace: 'stub-inline-worker' }, () => ({
      contents:
        "export default class MainThreadOnlyWorker {" +
        " constructor() { throw new Error('[ooxml] parser worker is not available in the VS Code extension (main-thread only)'); }" +
        " }",
      loader: 'js',
    }));
  },
};
