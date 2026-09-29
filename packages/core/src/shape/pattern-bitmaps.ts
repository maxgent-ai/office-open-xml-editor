// ECMA-376 Part 1 §20.1.10.51 / ST_PresetPatternVal specifies the 54 names,
// not their cell artwork. The cells below were measured from PowerPoint's
// direct PDF export of every name in two opposite foreground/background colours.
// Each slash-separated row has eight cells: # = foreground, . = background,
// + = the sRGB midpoint. The diagonal families use midpoint edge cells.
// PowerPoint embeds a 64×64 image made of uniform 8×8 pixel blocks in an
// 8 pt PDF tiling pattern: one logical cell = 1 pt. PowerPoint, Word and Excel
// PDF controls with independent foreground/background alpha (50/50 and
// 25/75%; Excel also 0/50 and 100/0%) add a matching soft mask whose diagonal
// cells have midpoint alpha. PDF colour conversion can change RGB by 1–2
// levels. Both opaque colour variants and a 6× PowerPoint PDF raster agree at
// all 64 cells for all 54 names. The PPTX painter
// keeps that grid at the slide origin through rotation, reflection and groups.

import { createAuxCanvas } from '../canvas/aux-canvas.js';

const PATTERN_BITMAPS: Record<string, string> = {
  pct5: '#......./......../......../......../....#.../......../......../........',
  pct10: '#......./......../....#.../......../#......./......../....#.../........',
  pct20: '#...#.../......../..#...#./......../#...#.../......../..#...#./........',
  pct25: '#...#.../..#...#./#...#.../..#...#./#...#.../..#...#./#...#.../..#...#.',
  pct30: '#.#.#.#./.#...#../#.#.#.#./...#...#/#.#.#.#./.#...#../#.#.#.#./...#...#',
  pct40: '#.#.#.#./.#.#.#.#/#.#.#.#./.#.#...#/#.#.#.#./.#.#.#.#/#.#.#.#./...#.#.#',
  pct50: '#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#',
  pct60: '###.###./.#.#.#.#/#.###.##/.#.#.#.#/###.###./.#.#.#.#/#.###.##/.#.#.#.#',
  pct70: '.###.###/##.###.#/.###.###/##.###.#/.###.###/##.###.#/.###.###/##.###.#',
  pct75: '.###.###/########/##.###.#/########/.###.###/########/##.###.#/########',
  pct80: '###.####/########/#######./########/###.####/########/#######./########',
  pct90: '########/########/########/####.###/########/########/########/.#######',
  horz: '########/......../......../......../......../......../......../........',
  vert: '#......./#......./#......./#......./#......./#......./#......./#.......',
  ltHorz: '########/......../......../......../########/......../......../........',
  ltVert: '#...#.../#...#.../#...#.../#...#.../#...#.../#...#.../#...#.../#...#...',
  dkHorz: '########/########/......../......../########/########/......../........',
  dkVert: '##..##../##..##../##..##../##..##../##..##../##..##../##..##../##..##..',
  narHorz: '########/......../########/......../########/......../########/........',
  narVert: '.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#/.#.#.#.#',
  dashHorz: '####..../......../......../......../....####/......../......../........',
  dashVert: '#......./#......./#......./#......./....#.../....#.../....#.../....#...',
  cross: '########/#......./#......./#......./#......./#......./#......./#.......',
  dnDiag: '#+.....+/+#+...../.+#+..../..+#+.../...+#+../....+#+./.....+#+/+.....+#',
  upDiag: '+.....+#/.....+#+/....+#+./...+#+../..+#+.../.+#+..../+#+...../#+.....+',
  ltDnDiag: '#...#.../.#...#../..#...#./...#...#/#...#.../.#...#../..#...#./...#...#',
  ltUpDiag: '...#...#/..#...#./.#...#../#...#.../...#...#/..#...#./.#...#../#...#...',
  dkDnDiag: '##..##../.##..##./..##..##/#..##..#/##..##../.##..##./..##..##/#..##..#',
  dkUpDiag: '..##..##/.##..##./##..##../#..##..#/..##..##/.##..##./##..##../#..##..#',
  wdDnDiag: '##.....#/###...../.###..../..###.../...###../....###./.....###/#.....##',
  wdUpDiag: '#.....##/.....###/....###./...###../..###.../.###..../###...../##.....#',
  dashDnDiag: '......../......../#...#.../.#...#../..#...#./...#...#/......../........',
  dashUpDiag: '......../......../...#...#/..#...#./.#...#../#...#.../......../........',
  diagCross: '#+....+#/+#+..+#+/.+#++#+./..+##+../..+##+../.+#++#+./+#+..+#+/#+....+#',
  smCheck: '#..##..#/.##..##./.##..##./#..##..#/#..##..#/.##..##./.##..##./#..##..#',
  lgCheck: '####..../####..../####..../####..../....####/....####/....####/....####',
  smGrid: '########/#...#.../#...#.../#...#.../########/#...#.../#...#.../#...#...',
  lgGrid: '########/#......./#......./#......./#......./#......./#......./#.......',
  dotGrid: '#.#.#.#./......../#......./......../#......./......../#......./........',
  smConfetti: '#......./....#.../.#....../......#./...#..../.......#/..#...../.....#..',
  lgConfetti: '#.##...#/..##..../......##/...##.##/##.##.../##....../....##../#...##.#',
  horzBrick: '########/#......./#......./#......./########/....#.../....#.../....#...',
  diagBrick: '.......#/......#./.....#../....#.../...##.../..#..#../.#....#./#......#',
  solidDmnd: '...#..../..###.../.#####../#######./.#####../..###.../...#..../........',
  openDmnd: '#.....#./.#...#../..#.#.../...#..../..#.#.../.#...#../#.....#./.......#',
  dotDmnd: '#......./......../..#...#./......../....#.../......../..#...#./........',
  plaid: '#.#.#.#./.#.#.#.#/#.#.#.#./.#.#.#.#/####..../####..../####..../####....',
  sphere: '.###.###/#...#..#/#...####/#...####/.###.###/#..##.../#####.../#####...',
  weave: '#...#.../.#.#.#../..#...#./.#...#.#/#...#.../...#.#../..#...#./.#.#...#',
  divot: '......../...#..../....#.../...#..../......../#......./.......#/#.......',
  shingle: '......##/#....#../.#..#.../..##..../....##../......#./.......#/.......#',
  wave: '......../...##.../..#..#.#/##....../......../...##.../..#..#.#/##......',
  trellis: '########/.##..##./########/#..##..#/########/.##..##./########/#..##..#',
  zigZag: '#......#/.#....#./..#..#../...##.../#......#/.#....#./..#..#../...##...',
};

/** Render each measured cell at PowerPoint PDF's native 8×8 sample resolution.
 * A one-pixel source cell would be interpolated across its whole interior
 * when a slide is enlarged. At small output sizes PDF and Canvas rasterizers
 * blend cell boundaries differently; the source tile stays unfiltered so its
 * interior colours and high-resolution phase remain exact. */
export function buildPatternBitmap(
  preset: string,
  fg: string,
  bg: string,
): HTMLCanvasElement | OffscreenCanvas | null {
  const encoded = PATTERN_BITMAPS[preset];
  if (typeof encoded !== 'string') return null;
  const tile = createAuxCanvas(64, 64);
  if (!tile) return null;
  const ctx = tile.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) return null;

  const bgCss = hexToCss(bg);
  const fgCss = hexToCss(fg);
  const midpointCss = blendHalfCss(fg, bg);
  const rows = encoded.split('/');
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      const cell = rows[y][x];
      // The Office PDF tile stores one RGBA value per cell. In controls with
      // independent foreground/background a:alpha (0/50, 25/75 and 50/50%),
      // the half-covered diagonal cells store each RGB and alpha channel at
      // the pair's midpoint. Painting a translucent cell over a prefilled
      // background would increase its alpha, so paint every cell exactly once.
      ctx.fillStyle = cell === '#' ? fgCss : cell === '+' ? midpointCss : bgCss;
      ctx.fillRect(x * 8, y * 8, 8, 8);
    }
  }
  return tile;
}

function hexToCss(hex: string): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const a = hex.length >= 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1;
  return `rgba(${r},${g},${b},${a})`;
}

function blendHalfCss(fg: string, bg: string): string {
  const channel = (hex: string, offset: number): number => parseInt(hex.slice(offset, offset + 2), 16);
  const r = Math.round((channel(fg, 0) + channel(bg, 0)) / 2);
  const g = Math.round((channel(fg, 2) + channel(bg, 2)) / 2);
  const b = Math.round((channel(fg, 4) + channel(bg, 4)) / 2);
  const alpha = (hex: string): number => hex.length >= 8 ? channel(hex, 6) / 255 : 1;
  return `rgba(${r},${g},${b},${(alpha(fg) + alpha(bg)) / 2})`;
}
