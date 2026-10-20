import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { expect, it } from 'vitest';
import { OoxmlError, type ModelSource } from '@silurus/ooxml-core';
import { MODEL_SOURCE_MODULE_PROTOCOL } from '@silurus/ooxml-core/internal/model-source';
import { resolveNodeSessionInput } from './model-source.js';

it('normalizes source open and lazy failures, preserving frozen methods and their receiver', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ooxml-source-rejection-'));
  const path = join(directory, 'source.mjs');
  writeFileSync(path, `
    const failure = 'MODEL_SOURCE_UNSUPPORTED:{"format":"custom","reason":"source-subset"}';
    export async function openModelSource(_bytes, config) {
      if (config.failOpen) throw failure;
      const archive = Object.freeze({ value: 7, read() { if (this.value !== 7) throw new Error('wrong receiver'); throw failure; } });
      return { archive, close() {} };
    }`);
  const source = (failOpen: boolean): ModelSource<'docx'> => ({
    target: 'docx', claim: () => true,
    beginLoad: () => ({ module: { protocol: MODEL_SOURCE_MODULE_PROTOCOL, target: 'docx',
      moduleUrl: pathToFileURL(path).href, config: { failOpen } }, release() {} }),
  });
  const open = (failOpen: boolean) => resolveNodeSessionInput(new Uint8Array(), 'docx',
    { source: source(failOpen) }, value => value as { read(): void });
  try {
    await expect(open(true)).rejects.toBeInstanceOf(OoxmlError);
    const result = await open(false);
    expect(result.kind).toBe('model-source');
    if (result.kind !== 'model-source') throw new Error('source was not selected');
    expect(() => result.opened.archive.read()).toThrow(expect.objectContaining({ code: 'unsupported-content',
      details: { format: 'custom', reason: 'source-subset' } }));
    result.opened.close();
  } finally { rmSync(directory, { recursive: true, force: true }); }
});
