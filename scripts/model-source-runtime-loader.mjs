// Node loader hook used only by the distribution regression check. Recording
// resolution catches a fire-and-forget import even when the caller never awaits
// it. The companion preload watches synchronous CommonJS loading.
import { appendFileSync } from 'node:fs';

export async function resolve(specifier, context, nextResolve) {
  if (/model-source|source-worker/.test(specifier) && !specifier.includes('/scripts/')) {
    appendFileSync(process.env.OOXML_SOURCE_REQUEST_LOG, `${specifier}\n`);
  }
  return nextResolve(specifier, context);
}

export async function load(url, context, nextLoad) {
  const result = await nextLoad(url, context);
  const source = typeof result.source === 'string' ? result.source
    : result.source ? Buffer.from(result.source).toString('utf8') : '';
  if (source.includes('ooxml-model-source-module/v1')
    || source.includes('model source view default')) {
    appendFileSync(process.env.OOXML_SOURCE_REQUEST_LOG, `${url}\n`);
  }
  return result;
}
