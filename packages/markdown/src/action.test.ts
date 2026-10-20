import { describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

// The GitHub Action converts files found by walking the repository. Their
// names are attacker-controlled (any contributor can commit a file), so a
// name must never reach a shell command line.
const ACTION = resolve(import.meta.dirname, '../action/index.js');

describe('ooxml-to-markdown GitHub Action', () => {
  it('passes repository file names to child processes without a shell', () => {
    const root = mkdtempSync(join(tmpdir(), 'ooxml-md-action-'));
    try {
      const marker = join(root, 'PWNED');
      for (const name of ['a$(touch PWNED).docx', 'b`touch PWNED`.docx', 'c";touch PWNED;".docx']) {
        writeFileSync(join(root, name), 'not a package');
      }
      // A symlink (here to a file outside the walked tree) is never followed.
      const outside = mkdtempSync(join(tmpdir(), 'ooxml-md-outside-'));
      writeFileSync(join(outside, 'secret.docx'), 'not a package');
      symlinkSync(join(outside, 'secret.docx'), join(root, 'link.docx'));
      const run = spawnSync(process.execPath, [ACTION], {
        cwd: root,
        env: { ...process.env, GITHUB_WORKSPACE: root, INPUT_FILES: `${root}/**/*.docx`, INPUT_COMMIT: 'false' },
        encoding: 'utf8',
      });
      expect(run.stdout).toContain('Matched 3 files');
      expect(existsSync(marker)).toBe(false);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  }, 60_000);

  it('never reads or writes through symlinks that leave the workspace', () => {
    const root = mkdtempSync(join(tmpdir(), 'ooxml-md-action-'));
    const outside = mkdtempSync(join(tmpdir(), 'ooxml-md-outside-'));
    try {
      const sentinel = join(outside, 'sentinel');
      writeFileSync(sentinel, 'keep');
      writeFileSync(join(outside, 'external.docx'), 'not a package');
      // An output path that is a symlink to a file outside the workspace.
      writeFileSync(join(root, 'report.docx'), 'not a package');
      symlinkSync(sentinel, join(root, 'report.md'));
      // A symlinked walk root that points outside the workspace.
      symlinkSync(outside, join(root, 'docs'));
      const run = (files: string) => spawnSync(process.execPath, [ACTION], {
        cwd: root,
        env: { ...process.env, GITHUB_WORKSPACE: root, INPUT_FILES: files, INPUT_COMMIT: 'false' },
        encoding: 'utf8',
      });
      const direct = run(join(root, 'report.docx'));
      expect(direct.stderr).toContain('Refusing to write');
      expect(readFileSync(sentinel, 'utf8')).toBe('keep');
      // A symlinked out-dir component must not create directories outside.
      const outDir = spawnSync(process.execPath, [ACTION], {
        cwd: root,
        env: {
          ...process.env,
          GITHUB_WORKSPACE: root,
          INPUT_FILES: join(root, 'report.docx'),
          'INPUT_OUT-DIR': join(root, 'docs', 'new', 'deep'),
          INPUT_COMMIT: 'false',
        },
        encoding: 'utf8',
      });
      expect(outDir.stderr).toContain('outside the workspace');
      expect(existsSync(join(outside, 'new'))).toBe(false);
      const walked = run(`${join(root, 'docs')}/**/*.docx`);
      expect(walked.stdout).toContain('Matched 0 files');
      expect(run(join(root, 'docs', 'external.docx')).stdout).toContain('Matched 0 files');
    } finally {
      rmSync(root, { recursive: true, force: true });
      rmSync(outside, { recursive: true, force: true });
    }
  }, 60_000);
});
