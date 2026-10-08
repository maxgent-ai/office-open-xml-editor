// Self-authored whole-shape path fills exercise the published PPTX viewer.
import { storedZip } from './chart-ex-packages.mjs';

export function pathGradientPptxBytes() {
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const shapes = ['rect', 'shape'].map((path, index) => `<p:sp>
    <p:nvSpPr><p:cNvPr id="${index + 2}" name="${path}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm><a:off x="${(10 + index * 240) * 9525}" y="190500"/>
      <a:ext cx="1905000" cy="1905000"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      <a:gradFill><a:gsLst>
        <a:gs pos="0"><a:srgbClr val="000000"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="808080"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
      </a:gsLst><a:path path="${path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
      <a:tileRect l="0" t="0" r="0" b="0"/></a:gradFill><a:ln><a:noFill/></a:ln>
    </p:spPr></p:sp>`).join('');
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
      <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
    </Types>`],
    ['_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`],
    ['ppt/presentation.xml', `<p:presentation xmlns:p="${p}" xmlns:r="${r}">
      <p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst>
      <p:sldSz cx="4572000" cy="2286000"/></p:presentation>`],
    ['ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdSlide" Type="${r}/slide" Target="slides/slide1.xml"/></Relationships>`],
    ['ppt/slides/slide1.xml', `<p:sld xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`],
  ]);
}
