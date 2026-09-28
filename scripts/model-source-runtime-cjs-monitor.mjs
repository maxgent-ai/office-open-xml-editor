// Preloaded in the distribution probe. ESM loader hooks do not observe
// createRequire(), so watch both CommonJS entry points before loading dist.
// This catches accidental eager loads; it does not defend against code that
// intentionally removes or evades the monitor.
import Module from 'node:module';
import { appendFileSync } from 'node:fs';

function record(specifier) {
  if (typeof specifier === 'string'
    && /model-source|source-worker/.test(specifier)
    && !specifier.includes('/scripts/')) {
    appendFileSync(process.env.OOXML_SOURCE_REQUEST_LOG, `${specifier}\n`);
  }
}

const originalLoad = Module._load;
Module._load = function monitoredLoad(specifier, parent, isMain) {
  record(specifier);
  return originalLoad.call(this, specifier, parent, isMain);
};

const originalRequire = Module.prototype.require;
Module.prototype.require = function monitoredRequire(specifier) {
  record(specifier);
  return originalRequire.call(this, specifier);
};
