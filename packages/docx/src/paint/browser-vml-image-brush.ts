import { imageNaturalSize } from '@silurus/ooxml-core';
import type { DeepReadonly, ImagePaintResourceDescriptor } from '../layout/types.js';
import type { VmlImageBrush } from './vml-image-brush.js';

/** Resource acquisition only: resolve native pixel size once after lazy decode,
 * before paint. Part 4 §19.1.2.5 defines size, origin, position and aspect;
 * native image pixels use the shared 96dpi image-coordinate convention. */
export function acquireVmlImageBrush(
  descriptor: DeepReadonly<ImagePaintResourceDescriptor>, image: CanvasImageSource,
): Readonly<VmlImageBrush> {
  const fill = descriptor.vmlFill;
  if (!fill) throw new Error('Missing VML image brush acquisition facts');
  const natural = imageNaturalSize(image);
  const box = descriptor.intrinsicSize;
  if (!(natural.w > 0 && natural.h > 0 && box.widthPt > 0 && box.heightPt > 0)) {
    throw new RangeError('VML image brush requires positive image and host dimensions');
  }
  let width = fill.size
    ? fill.size[0].value * (fill.size[0].fraction ? box.widthPt : 1)
    : fill.repeat ? natural.w * 72 / 96 : box.widthPt;
  let height = fill.size
    ? fill.size[1].value * (fill.size[1].fraction ? box.heightPt : 1)
    : fill.repeat ? natural.h * 72 / 96 : box.heightPt;
  if (fill.aspect === 'atleast' || fill.aspect === 'atmost') {
    const ratio = (fill.aspect === 'atleast' ? Math.max : Math.min)(width / natural.w, height / natural.h);
    width = natural.w * ratio; height = natural.h * ratio;
  }
  if (![width, height].every(n => Number.isFinite(n) && n > 0)) {
    throw new RangeError('VML image brush size must be finite and positive');
  }
  return Object.freeze({ kind: 'vml-image-brush', image, repeat: fill.repeat,
    width: width / box.widthPt, height: height / box.heightPt,
    x: fill.position[0] - fill.origin[0] * width / box.widthPt,
    y: fill.position[1] - fill.origin[1] * height / box.heightPt,
    pixelWidth: natural.w, pixelHeight: natural.h,
  });
}
