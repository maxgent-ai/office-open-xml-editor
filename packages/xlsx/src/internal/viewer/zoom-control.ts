import { ListenerScope } from './listener-scope.js';

// Magnetic dead zone around the slider's 100% center notch. This is measured in
// slider-position units rather than scale points because the two halves map
// different scale spans (zoomMin→1 and 1→zoomMax); a position radius therefore
// gives the thumb the same physical attraction distance from either direction.
const ZOOM_SLIDER_100_SNAP_RADIUS = 2;

/** Map a slider position [0,100] to a scale factor. 50 → 1.0 (100%), with a
 *  separate linear segment on each side so the center is always 100%. */
export function zoomPosToScale(pos: number, min: number, max: number): number {
  return pos <= 50
    ? min + (pos / 50) * (1 - min)
    : 1 + ((pos - 50) / 50) * (max - 1);
}

/** Inverse of {@link zoomPosToScale}: scale factor → slider position [0,100]. */
export function zoomScaleToPos(scale: number, min: number, max: number): number {
  const clamped = Math.min(max, Math.max(min, scale));
  return clamped <= 1
    ? ((clamped - min) / (1 - min)) * 50
    : 50 + ((clamped - 1) / (max - 1)) * 50;
}

/** Zoom actions the footer control drives on its engine. */
export interface ZoomControlHost {
  setScale(scale: number): void;
  zoomIn(): void;
  zoomOut(): void;
}

/**
 * Excel-style zoom control pinned to the footer's logical end:
 * `−  [────slider────]  +  100%`. Live-updates the cell scale on input; the
 * engine reports every resolved scale back through {@link sync}.
 */
export class ZoomControl {
  readonly element: HTMLDivElement;
  readonly slider: HTMLInputElement;
  readonly label: HTMLSpanElement;
  private readonly listeners = new ListenerScope();

  constructor(
    ownerDocument: Document,
    private readonly host: ZoomControlHost,
    scale: number,
    zoomMin: number,
    zoomMax: number,
  ) {

    const wrap = ownerDocument.createElement('div');
    wrap.style.cssText =
      `display:flex;align-items:center;flex-shrink:0;gap:2px;` +
      `padding:0 10px;height:100%;` +
      `color:var(--ooxml-xlsx-chrome-text-muted,#555);font-size:12px;user-select:none;`;

    // The steppers walk the shared IX9 zoom ladder (ZOOM_STEP_LADDER via
    // zoomIn/zoomOut) so the built-in chrome and a host's own buttons wired to
    // the ZoomableViewer contract land on identical scales (issue #842).
    // Pre-IX9 these stepped ±0.1 linearly.
    const mkBtn = (glyph: string, label: string, step: () => void): HTMLButtonElement => {
      const b = ownerDocument.createElement('button');
      b.type = 'button';
      b.textContent = glyph;
      b.setAttribute('aria-label', label);
      b.title = label;
      b.style.cssText =
        `width:18px;height:18px;padding:0;border:none;background:transparent;` +
        `color:var(--ooxml-xlsx-chrome-text-muted,#555);` +
        `font-size:14px;line-height:1;cursor:pointer;border-radius:3px;`;
      this.listeners.on(b, 'click', step);
      return b;
    };

    // The slider works in "position" units [0,100]; 50 is dead-center and maps
    // to 100% so each half is its own linear segment (zoomMin→1 on the left,
    // 1→zoomMax on the right), mirroring Excel's status-bar zoom where 100% sits
    // in the middle even though the range (10%–400%) is asymmetric.
    const slider = ownerDocument.createElement('input');
    slider.type = 'range';
    slider.min = '0';
    slider.max = '100';
    slider.step = 'any';
    slider.value = String(zoomScaleToPos(scale, zoomMin, zoomMax));
    slider.setAttribute('aria-label', 'Zoom');
    slider.title = 'Zoom';
    slider.classList.add('xlsx-zoom-slider');
    slider.style.cssText = `width:90px;cursor:pointer;`;
    this.listeners.on(slider, 'input', () => {
      const rawPos = Number(slider.value);
      const pos = Math.abs(rawPos - 50) <= ZOOM_SLIDER_100_SNAP_RADIUS ? 50 : rawPos;
      // Move the thumb as well as the scale. setScale may otherwise return early
      // when the viewer is already at 100%, leaving the thumb beside the notch.
      if (pos === 50) slider.value = '50';
      this.host.setScale(zoomPosToScale(pos, zoomMin, zoomMax));
    });

    const label = ownerDocument.createElement('span');
    label.textContent = `${Math.round(scale * 100)}%`;
    label.style.cssText = `min-width:42px;margin-left:6px;text-align:right;font-variant-numeric:tabular-nums;`;

    wrap.appendChild(mkBtn('−', 'Zoom out', () => this.host.zoomOut()));
    wrap.appendChild(slider);
    wrap.appendChild(mkBtn('+', 'Zoom in', () => this.host.zoomIn()));
    wrap.appendChild(label);

    this.element = wrap;
    this.slider = slider;
    this.label = label;
  }

  /** Reflect a resolved scale (`percent` = its whole-percent label). */
  sync(scale: number, percent: number, zoomMin: number, zoomMax: number): void {
    this.slider.value = String(zoomScaleToPos(scale, zoomMin, zoomMax));
    this.label.textContent = `${percent}%`;
  }

  destroy(): void {
    this.listeners.dispose();
  }
}
