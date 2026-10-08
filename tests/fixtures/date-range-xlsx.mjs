// Minimal self-authored #1710 regression: one bad date plus a healthy sheet.
import { storedZip } from './chart-ex-packages.mjs';

export function dateRangeXlsxBytes() {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      ${[1, 2].map(n => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}
    </Types>`],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Dates" sheetId="1" r:id="s1"/><sheet name="Other" sheetId="2" r:id="s2"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}">
      <Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>
      <Relationship Id="s2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/>
      <Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/>
    </Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}">
      <fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts>
      <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
      <borders count="1"><border/></borders>
      <cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs>
    </styleSheet>`],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="1" width="24" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>12345678901</v></c></row></sheetData></worksheet>`],
    ['xl/worksheets/sheet2.xml', `<worksheet xmlns="${main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Healthy sheet</t></is></c></row></sheetData></worksheet>`],
  ]);
}
