// Browser bundle entry for docx-watermark-phase.spec.ts: the page painter plus
// the layout-time geometry resolver whose result every retained DrawingML
// paint plan carries.
export { paintLayoutPage } from '../../../packages/docx/src/paint/canvas-page.js';
export { resolveDrawingMLGeometry } from '../../../packages/core/src/shape/drawingml-geometry.js';
