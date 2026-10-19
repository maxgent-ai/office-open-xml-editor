export { mathToMathML } from './mathml';
// Only the light, engine-free helpers are re-exported here (and from the package
// index). The heavy MathJax engine (`loadMathJax`, `mathMLToSvg`) lives in
// `./engine` and is published separately as `@silurus/ooxml/math` so it stays
// out of the docx/pptx bundles unless the consumer opts in. See `MathRenderer`.
export {
  svgExtents,
  recolorSvg,
  type MathSvg,
  type MathRenderer,
} from './mathjax';
export {
  rasterizeMathSvg,
  sizeMathSvgForRaster,
  tintMathRaster,
  cachedTintedMathRaster,
  type RasterizedMathSvg,
} from './raster';
