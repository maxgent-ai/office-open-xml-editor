// Self-authored built-in time-format fixture; intentionally has no numFmts.
import { storedZip } from './chart-ex-packages.mjs';

export function builtinTimeXlsxBytes() {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const ids = [20, 21, 45, 46, 47];
  const values = [0.75, 45200.75, 62.34 / 86400, 59.96 / 86400];
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Times" sheetId="1" r:id="s1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="5">${ids.map(id => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>`).join('')}</cellXfs></styleSheet>`],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><cols><col min="1" max="4" width="25" customWidth="1"/></cols><sheetData>${ids.map((_, i) => `<row r="${i + 1}">${values.map((v, j) => `<c r="${'ABCD'[j]}${i + 1}" s="${i}"><v>${v}</v></c>`).join('')}</row>`).join('')}</sheetData></worksheet>`],
  ]);
}
