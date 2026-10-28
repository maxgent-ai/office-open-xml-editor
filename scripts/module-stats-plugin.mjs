// Build-only diagnostics for the model-source dispatch-cost check
// (scripts/check-model-source-lazy-graph.mjs). When OOXML_MODULE_STATS names
// a file, the build writes each emitted chunk's static imports and the
// rendered code of every module it retains (before minification). Ordinary
// builds write nothing.
import { writeFileSync } from 'node:fs';
import { relative } from 'node:path';

export function moduleStatsPlugin(root = process.cwd()) {
  return {
    name: 'ooxml-module-stats',
    generateBundle(_options, bundle) {
      const path = process.env.OOXML_MODULE_STATS;
      if (!path) return;
      const chunks = {};
      for (const output of Object.values(bundle)) {
        if (output.type !== 'chunk') continue;
        const modules = {};
        for (const [id, module] of Object.entries(output.modules)) {
          if (module.renderedLength > 0) modules[relative(root, id).replaceAll('\\', '/')] = module.code ?? '';
        }
        chunks[output.fileName] = { imports: [...output.imports], modules };
      }
      writeFileSync(path, JSON.stringify({ chunks }));
    },
  };
}
