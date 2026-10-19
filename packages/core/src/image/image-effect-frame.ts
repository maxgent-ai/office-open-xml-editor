import type { RgbaBuffer } from './duotone.js';

/** The authored image frame within a grown effect grid, in pixels. This is
 * decoded-image geometry, never a retained layout or parser mutation. */
export interface EffectFrame { x: number; y: number; width: number; height: number }
export interface EffectBuffer extends RgbaBuffer { frame?: EffectFrame }
const frames = new WeakMap<object, EffectFrame>();
export function imageEffectFrame(image: object): EffectFrame | undefined { return frames.get(image); }
export function setImageEffectFrame(image: object, frame: EffectFrame | undefined): void {
  if (frame) frames.set(image, frame);
}


/** Whether transformed pixels extend outside their authored picture frame. */
export function hasImageEffectOutset(image: object): boolean {
  const frame = frames.get(image);
  if (!frame) return false;
  const size = image as { width: number; height: number };
  return frame.x > 0 || frame.y > 0 || size.width > frame.x + frame.width || size.height > frame.y + frame.height;
}

/** How many whole tile periods an effect outset (blur grow, glow, shadow)
 * reaches beyond its authored frame on each axis. A repeating painter must
 * draw that many extra neighbours per side so their outsets wrap into the
 * covered area (§20.1.8.58 repeats the frame; the outset overlaps). A cropped
 * picture paints no outset (see drawImageCropped), so it needs none. */
export function imageEffectOutsetPeriods(image: object, cropped: boolean): { x: number; y: number } {
  const frame = frames.get(image);
  if (!frame || cropped || !(frame.width > 0) || !(frame.height > 0)) return { x: 0, y: 0 };
  const size = image as { width: number; height: number };
  return {
    x: Math.ceil(Math.max(0, frame.x, size.width - frame.x - frame.width) / frame.width),
    y: Math.ceil(Math.max(0, frame.y, size.height - frame.y - frame.height) / frame.height),
  };
}
