import { storedZip } from '../../../../tests/fixtures/chart-ex-packages.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = entries => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const blip = (id, mode = '<a:stretch><a:fillRect/></a:stretch>') => `<a:blipFill><a:blip r:embed="${id}"/>${mode}</a:blipFill>`;
const ref = '<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>';
const solid = '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>';
const pattern = '<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>';

// Self-authored normative fill fixtures, with deliberately colliding rIds in
// theme, style and slide parts. No text: row measurement is outside this test.
export function tableImageFillBytes(kind) {
  const themeBackground = kind === 'theme-background';
  const directBackground = kind === 'direct-background' || kind === 'direct-table' || kind === 'pattern-background';
  const cellFill = kind === 'direct-cell' ? blip('sameId') : kind === 'pattern-cell' ? pattern : '';
  const role = kind === 'theme-band' ? ref : kind === 'direct-band' ? `<a:fill>${blip('sameId')}</a:fill>` : `<a:fill>${solid}</a:fill>`;
  const tableStyleId = themeBackground ? '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}' : '{11111111-1111-1111-1111-111111111111}';
  const entries = [
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="svg" ContentType="image/svg+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/></Types>`],
    ['_rels/.rels', rels([['pres', 'officeDocument', 'ppt/presentation.xml']])],
    ['ppt/presentation.xml', `<p:presentation xmlns:p="${P}" xmlns:r="${R}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`],
    ['ppt/_rels/presentation.xml.rels', rels([['slide', 'slide', 'slides/slide1.xml'], ['theme', 'theme', 'theme/theme1.xml'], ['styles', 'tableStyles', 'tableStyles.xml']])],
    ['ppt/theme/theme1.xml', `<a:theme xmlns:a="${A}" xmlns:r="${R}" name="test"><a:themeElements><a:clrScheme name="test">${['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].map(name => `<a:${name}><a:srgbClr val="FF0000"/></a:${name}>`).join('')}</a:clrScheme><a:fontScheme name="test"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="test"><a:fillStyleLst>${solid}${blip('sameId')}${pattern}</a:fillStyleLst><a:lnStyleLst>${'<a:ln/>'.repeat(3)}</a:lnStyleLst><a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${solid.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`],
    ['ppt/theme/_rels/theme1.xml.rels', rels([['sameId', 'image', '../media/theme.svg']])],
    ['ppt/tableStyles.xml', `<a:tblStyleLst xmlns:a="${A}" xmlns:r="${R}" def="${tableStyleId}"><a:tblStyle styleId="{11111111-1111-1111-1111-111111111111}" styleName="test">${directBackground ? `<a:tblBg><a:fill>${kind === 'pattern-background' ? pattern : blip('sameId')}</a:fill></a:tblBg>` : ''}<a:wholeTbl><a:tcStyle><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle>${role}</a:tcStyle></a:band1H></a:tblStyle></a:tblStyleLst>`],
    ['ppt/_rels/tableStyles.xml.rels', rels([['sameId', 'image', 'media/style.svg']])],
    ['ppt/slides/slide1.xml', `<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table"><a:tbl><a:tblPr bandRow="1">${kind === 'direct-table' ? blip('sameId') : ''}<a:tableStyleId>${tableStyleId}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="9144000"/></a:tblGrid>${[0,1].map(() => `<a:tr h="3429000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr>${cellFill}</a:tcPr></a:tc></a:tr>`).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>`],
    ['ppt/slides/_rels/slide1.xml.rels', rels([['sameId', 'image', '../media/cell.svg']])],
  ];
  for (const [name, color] of [['theme', '0000FF'], ['style', '00FF00'], ['cell', 'FFFF00']]) {
    entries.push([`ppt/media/${name}.svg`, `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32"><rect width="32" height="32" fill="#${color}"/></svg>`]);
  }
  return storedZip(entries);
}
