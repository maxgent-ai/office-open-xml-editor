/**
 * Stateful-archive-handle benchmark (Phase 2 D2 + D3).
 *
 * `bench-parse.mjs` measures a single `parse` across the WASM boundary. This
 * script measures the *repeated* work the viewer actually does after a parse:
 *
 *   - xlsx: parse the workbook index, then parse EVERY sheet (sheet switching).
 *   - docx / pptx: parse the document, then extract EVERY embedded image.
 *
 * It runs each workload two ways and reports both so the win is visible:
 *
 *   - "free"   : the pre-handle path — call the free functions repeatedly, each
 *                re-copying the whole file into WASM and re-opening the ZIP (and,
 *                for xlsx, re-parsing workbook.xml / sharedStrings / theme on
 *                every sheet). This is what the code paid before this change.
 *   - "handle" : the new path — build one `{Xlsx,Pptx,Docx}Archive`, then call
 *                its methods. The bytes are copied + the ZIP opened ONCE; xlsx
 *                additionally parses the shared workbook parts once and reuses
 *                them for every sheet (the D3 win).
 *
 * The handle classes only exist on the post-change WASM. When they are absent
 * (running against origin/main WASM for a before/after comparison) the "handle"
 * mode is reported as "n/a" and only "free" is measured — so a single script,
 * pointed at the before WASM and then the after WASM, yields the full picture:
 *
 *   BEFORE (origin/main):  handle=n/a,  free=<baseline>
 *   AFTER  (this branch):  handle=<fast>, free≈<baseline>   (free path unchanged)
 *
 * Usage:
 *   node packages/node/src/bench-handle.mjs <file> [iterations] [--wasm-dir <dir>]
 *
 *   --wasm-dir <dir>   Load *_parser.js / *_parser_bg.wasm from <dir> instead of
 *                      the package's src/wasm. Point it at a saved "before" build
 *                      to measure the baseline. <dir> may contain the files
 *                      directly or in a per-format subdir (docx/ pptx/ xlsx/).
 *
 * Requires freshly built WASM (`pnpm build:wasm`) for the "after" run.
 */
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve, extname, basename, join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));
const WARMUP = 1;
const DEFAULT_ITERS = 5;
const decoder = new TextDecoder();

/** Load a wasm-pack `--target web` module and synchronously init it from disk. */
function loadModule(jsPath, wasmPath) {
  const bytes = readFileSync(wasmPath);
  const wasmModule = new WebAssembly.Module(bytes);
  return import(jsPath).then((mod) => {
    mod.initSync({ module: wasmModule });
    return mod;
  });
}

/** Resolve the JS + WASM paths for a format, honoring `--wasm-dir`. */
function wasmPaths(fmt, wasmDir) {
  const stem = `${fmt}_parser`;
  if (wasmDir) {
    // Accept either "<dir>/<fmt>/<stem>.*" or "<dir>/<stem>.*".
    const sub = join(wasmDir, fmt, `${stem}.js`);
    const base = join(wasmDir, `${stem}.js`);
    const jsPath = existsSync(sub) ? sub : base;
    const wasmPath = jsPath.replace(/\.js$/, '_bg.wasm');
    return { jsPath, wasmPath };
  }
  const jsPath = resolve(HERE, `../../${fmt}/src/wasm/${stem}.js`);
  const wasmPath = resolve(HERE, `../../${fmt}/src/wasm/${stem}_bg.wasm`);
  return { jsPath, wasmPath };
}

function stats(samples) {
  const s = [...samples].sort((a, b) => a - b);
  const n = s.length;
  const median = n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
  return { min: s[0], median, mean: s.reduce((a, x) => a + x, 0) / n };
}

/** Time `fn` `iters` times after `WARMUP` untimed runs; returns per-run ms. */
function timeIt(fn, iters) {
  for (let i = 0; i < WARMUP; i++) fn();
  const samples = [];
  for (let i = 0; i < iters; i++) {
    const t0 = performance.now();
    const sink = fn();
    const t1 = performance.now();
    if (sink == null) throw new Error('workload produced null');
    samples.push(t1 - t0);
  }
  return stats(samples);
}

// ── xlsx: parse workbook + parse every sheet ────────────────────────────────

function xlsxWorkloads(mod, bytes) {
  const wbFree = () => JSON.parse(decoder.decode(mod.parse_xlsx(bytes, undefined)));
  const sheets = wbFree().workbook.sheets;

  const free = () => {
    // Re-parse the index (as the worker's `parse` does) + every sheet.
    let touched = 0;
    mod.parse_xlsx(bytes, undefined);
    for (let i = 0; i < sheets.length; i++) {
      const ws = mod.parse_sheet(bytes, i, sheets[i].name, undefined);
      touched += ws.length;
    }
    return touched;
  };

  const handle = mod.XlsxArchive
    ? () => {
        let touched = 0;
        const ar = new mod.XlsxArchive(bytes, undefined);
        try {
          ar.parse();
          for (let i = 0; i < sheets.length; i++) {
            touched += ar.parse_sheet(i, sheets[i].name).length;
          }
        } finally {
          ar.free();
        }
        return touched;
      }
    : null;

  return { free, handle, unit: `${sheets.length} sheets`, count: sheets.length };
}

// ── docx / pptx: parse + extract every image ────────────────────────────────

/** Collect every embedded media/image zip path from a parsed model. */
function collectImagePaths(model, fmt) {
  const paths = new Set();
  const prefix = fmt === 'pptx' ? 'ppt/media/' : 'word/media/';
  const walk = (v) => {
    if (v == null) return;
    if (typeof v === 'string') {
      if (v.startsWith(prefix)) paths.add(v);
      return;
    }
    if (Array.isArray(v)) {
      for (const x of v) walk(x);
      return;
    }
    if (typeof v === 'object') for (const k of Object.keys(v)) walk(v[k]);
  };
  walk(model);
  return [...paths];
}

function docParseWorkloads(mod, bytes, fmt) {
  const parseFn = fmt === 'pptx' ? mod.parse_pptx : mod.parse_docx;
  const Handle = fmt === 'pptx' ? mod.PptxArchive : mod.DocxArchive;
  const model = JSON.parse(decoder.decode(parseFn(bytes, undefined)));
  const imgs = collectImagePaths(model, fmt);

  const free = () => {
    let touched = 0;
    parseFn(bytes, undefined);
    for (const p of imgs) touched += mod.extract_image(bytes, p, undefined).length;
    return touched + 1;
  };

  const handle = Handle
    ? () => {
        let touched = 0;
        const ar = new Handle(bytes, undefined);
        try {
          ar.parse();
          for (const p of imgs) touched += ar.extract_image(p).length;
        } finally {
          ar.free();
        }
        return touched + 1;
      }
    : null;

  return { free, handle, unit: `${imgs.length} images`, count: imgs.length };
}

function fmtRow(label, r) {
  if (!r) return `  ${label.padEnd(8)}: n/a (handle class not in this WASM)`;
  const f = (x) => x.toFixed(2).padStart(9);
  return `  ${label.padEnd(8)}: min${f(r.min)}  median${f(r.median)}  mean${f(r.mean)} ms`;
}

async function main() {
  const argv = process.argv.slice(2);
  let wasmDir = null;
  const di = argv.indexOf('--wasm-dir');
  if (di !== -1) {
    wasmDir = resolve(process.cwd(), argv[di + 1]);
    argv.splice(di, 2);
  }
  const [file, itersArg] = argv;
  if (!file) {
    console.error(
      'usage: node packages/node/src/bench-handle.mjs <file> [iterations] [--wasm-dir <dir>]',
    );
    process.exit(1);
  }
  const iters = itersArg ? Number(itersArg) : DEFAULT_ITERS;
  const fmt = extname(file).toLowerCase().slice(1);
  if (!['xlsx', 'pptx', 'docx'].includes(fmt)) {
    throw new Error(`unsupported extension: .${fmt} (expected .xlsx/.pptx/.docx)`);
  }
  const bytes = readFileSync(resolve(process.cwd(), file));

  const { jsPath, wasmPath } = wasmPaths(fmt, wasmDir);
  const mod = await loadModule(jsPath, wasmPath);

  const w = fmt === 'xlsx' ? xlsxWorkloads(mod, bytes) : docParseWorkloads(mod, bytes, fmt);
  const source = wasmDir ? `before (${basename(wasmDir)})` : 'after (src/wasm)';

  const freeStats = timeIt(w.free, iters);
  const handleStats = w.handle ? timeIt(w.handle, iters) : null;

  const speedup = handleStats ? (freeStats.median / handleStats.median).toFixed(2) : 'n/a';

  console.log(`file      : ${basename(file)} (${(bytes.length / 1_000_000).toFixed(2)} MB)`);
  console.log(`workload  : parse + ${w.unit}   [${fmt}]   wasm=${source}`);
  console.log(`iters     : warmup=${WARMUP} iters=${iters}`);
  console.log(fmtRow('free', freeStats));
  console.log(fmtRow('handle', handleStats));
  console.log(`  speedup : ${speedup}x  (free median / handle median)`);
  // Machine-readable line for scripted collection.
  const h = handleStats ? handleStats.median.toFixed(2) : 'na';
  console.log(
    `RESULT\t${basename(file)}\t${fmt}\t${source}\t${w.count}\tfree=${freeStats.median.toFixed(2)}\thandle=${h}\tspeedup=${speedup}`,
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
