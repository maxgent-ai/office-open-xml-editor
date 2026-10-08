import { CanvasLoadingIndicator } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';

/** Build the shared progressive-slide loading surface used by both PPTX viewers. */
export function createPptxLoadingIndicator(wrapper: HTMLElement): CanvasLoadingIndicator {
  const ownerDocument = wrapper.ownerDocument ?? document;
  const indicator = new CanvasLoadingIndicator(wrapper, 'Loading slide');
  const layer = indicator.layer;
  layer.style.backdropFilter = 'blur(2px)';

  const circle = ownerDocument.createElement('span');
  circle.className = 'ooxml-pptx-progress-circle';
  circle.style.cssText = [
    'width:34px',
    'height:34px',
    'box-sizing:border-box',
    'border-radius:50%',
    'border:3px solid var(--border-bright, rgba(100,116,139,0.28))',
    'border-top-color:var(--signal, #12bfd8)',
    'box-shadow:0 2px 10px rgba(15,23,42,0.08)',
  ].join(';');
  circle.setAttribute('aria-hidden', 'true');
  layer.appendChild(circle);

  const reducedMotion = ownerDocument.defaultView
    ?.matchMedia?.('(prefers-reduced-motion: reduce)').matches ?? false;
  if (!reducedMotion && typeof circle.animate === 'function') {
    circle.animate(
      [{ transform: 'rotate(0deg)' }, { transform: 'rotate(360deg)' }],
      { duration: 800, iterations: Infinity, easing: 'linear' },
    );
  }

  return indicator;
}
