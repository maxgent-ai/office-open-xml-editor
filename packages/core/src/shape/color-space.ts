// DrawingML HSL on normalized sRGB, matching ooxml-common color::rgb_to_hls /
// hls_to_rgb (ECMA-376 Part 1 §20.1.2.3). H is in degrees, S/L in [0,1],
// RGB in [0,255]. Shared by raster effects and theme-derived chart colors.
export function rgbToHsl([red, green, blue]: readonly number[], out: [number, number, number] = [0, 0, 0]): [number, number, number] {
  const r = red / 255;
  const g = green / 255;
  const b = blue / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  const delta = max - min;
  if (delta === 0) { out[0] = 0; out[1] = 0; out[2] = lightness; return out; }
  const saturation = delta / (1 - Math.abs(2 * lightness - 1));
  const rawHue = max === r
    ? ((g - b) / delta) % 6
    : max === g
      ? (b - r) / delta + 2
      : (r - g) / delta + 4;
  out[0] = ((rawHue * 60) + 360) % 360; out[1] = saturation; out[2] = lightness; return out;
}

export function hslToRgb(hue: number, saturation: number, lightness: number, out: [number, number, number] = [0, 0, 0]): [number, number, number] {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const segment = ((hue % 360) + 360) % 360 / 60;
  const secondary = chroma * (1 - Math.abs(segment % 2 - 1));
  const [r1, g1, b1] = segment < 1 ? [chroma, secondary, 0]
    : segment < 2 ? [secondary, chroma, 0]
      : segment < 3 ? [0, chroma, secondary]
        : segment < 4 ? [0, secondary, chroma]
          : segment < 5 ? [secondary, 0, chroma]
            : [chroma, 0, secondary];
  const match = lightness - chroma / 2;
  out[0] = (r1 + match) * 255; out[1] = (g1 + match) * 255; out[2] = (b1 + match) * 255; return out;
}


/** Signed DrawingML channel interpolation; tint/shade and duotone share
 * this operation, with clamping at the byte boundary, never empirical gains. */
export function interpolateColorChannel(from: number, to: number, amount: number): number {
  return Math.round(Math.max(0, Math.min(255, from + (to - from) * amount)));
}
