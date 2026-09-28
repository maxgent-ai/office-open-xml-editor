#!/usr/bin/env node
// CI regression proof for accidental eager source loading, not a sandbox for
// deliberately adversarial code. Monitor both ESM and synchronous CommonJS
// requests because the latter bypass Node's asynchronous ESM loader hook.
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export function runSourceRequestProbe({
  script = './scripts/model-source-runtime-probe.mjs',
  cwd = process.cwd(),
} = {}) {
  const temp = mkdtempSync(join(process.platform === 'darwin' ? '/private/tmp' : tmpdir(), 'ooxml-source-requests-'));
  const log = join(temp, 'requests.log');
  try {
    writeFileSync(log, '');
    const result = spawnSync(process.execPath, [
      '--loader', './scripts/model-source-runtime-loader.mjs',
      '--import', './scripts/model-source-runtime-cjs-monitor.mjs',
      script,
    ], {
      cwd, encoding: 'utf8', timeout: 120_000,
      env: { ...process.env, OOXML_SOURCE_REQUEST_LOG: log },
    });
    if (result.status !== 0) {
      throw new Error(`Built Node OOXML probe failed: ${result.stderr || result.stdout}`);
    }
    const requests = readFileSync(log, 'utf8').trim();
    if (requests) throw new Error(`Default OOXML load requested optional source code:\n${requests}`);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  runSourceRequestProbe();
  console.log('Built Node OOXML loads requested no model-source modules.');
}
