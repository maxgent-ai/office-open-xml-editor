// Self-authored fixture: columns A:C frozen, unwrapped text overflowing across
// the freeze line in both directions (issue #1973).
import { storedZip } from './chart-ex-packages.mjs';

export function frozenOverflowXlsxBytes() {
  const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const text = (ref, value, style = 0) =>
    `<c r="${ref}" s="${style}" t="inlineStr"><is><t>${value}</t></is></c>`;
  const rows = [
    `<row r="1">${text('A1', 'A1')}${text('D1', 'D1')}</row>`,
    // Left-aligned in a frozen column: spills right into the scrollable pane.
    `<row r="4">${text('B4', 'LEFT-B4 aaaa|bbbb|cccc|dddd|eeee|ffff|gggg|hhhh|iiii|jjjj|kkkk|llll|mmmm|nnnn|oooo|pppp|END')}</row>`,
    // Right-aligned in a scrollable column: spills left into the frozen pane.
    `<row r="6">${text('E6', 'RIGHT-E6 rrrr|ssss|tttt|uuuu|vvvv|END', 1)}</row>`,
    // Centred on the freeze line's left side: spills both ways.
    `<row r="8">${text('C8', 'CENTER-C8 1111|2222|3333|4444|5555|6666|END', 2)}</row>`,
  ].join('');
  return storedZip([
    ['[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
    ['_rels/.rels', `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Frozen" sheetId="1" r:id="s1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${pkg}"><Relationship Id="s1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="styles" Type="${rel}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${main}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="right"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs></styleSheet>`],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${main}"><sheetViews><sheetView workbookViewId="0"><pane xSplit="3" ySplit="1" topLeftCell="D2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetData>${rows}</sheetData></worksheet>`],
  ]);
}
