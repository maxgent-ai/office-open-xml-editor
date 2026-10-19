export interface VmlImageBrush {
  readonly kind: 'vml-image-brush';
  readonly image: CanvasImageSource;
  readonly repeat: boolean;
  readonly x: number; readonly y: number;
  readonly width: number; readonly height: number;
  readonly pixelWidth: number; readonly pixelHeight: number;
}
export function isVmlImageBrush(handle: unknown): handle is Readonly<VmlImageBrush> {
  return typeof handle === 'object' && handle !== null && 'kind' in handle && handle.kind === 'vml-image-brush';
}
