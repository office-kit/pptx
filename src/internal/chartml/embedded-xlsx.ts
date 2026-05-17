// Minimal xlsx writer scoped to chart-data needs.
//
// PowerPoint requires an embedded workbook to back every chart so the
// "Edit data" affordance works. We don't need a real spreadsheet — just
// a single sheet with the chart's category labels and series values.
// Everything beyond that (styles, calculated columns, defined names,
// shared strings, charts within the xlsx itself) is deliberately
// excluded. The plan calls this out as a scope-creep risk; the file
// stays under `internal/chartml/` and is never exposed publicly.
//
// Schema (ECMA-376 Part 1 §18 SpreadsheetML, simplified):
//
//   /[Content_Types].xml      defaults + overrides
//   /_rels/.rels              → /xl/workbook.xml
//   /xl/workbook.xml          one <sheets>/<sheet> entry
//   /xl/_rels/workbook.xml.rels → /xl/worksheets/sheet1.xml
//   /xl/worksheets/sheet1.xml the data, using inline strings to skip
//                             the shared-strings table

import { writeZip } from '../opc/zip.ts';

const TEXT_ENCODER = new TextEncoder();
const encode = (s: string): Uint8Array => TEXT_ENCODER.encode(s);

const xmlEscape = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const colLetter = (col: number): string => {
  // 0-based to A, B, …, Z, AA, AB, … (Excel-style).
  let n = col;
  let out = '';
  while (true) {
    out = String.fromCharCode(65 + (n % 26)) + out;
    if (n < 26) return out;
    n = Math.floor(n / 26) - 1;
  }
};

const cellRef = (row: number, col: number): string => `${colLetter(col)}${row + 1}`;

/** A row in the embedded sheet — string in column 0, numeric values after. */
export interface DataRow {
  readonly label: string;
  readonly values: ReadonlyArray<number | null>;
}

/**
 * Builds the bytes of a fresh xlsx whose only sheet is laid out as
 *
 *   |       | seriesNames[0] | seriesNames[1] | ...
 *   |-------|----------------|----------------|----
 *   | row0  | row0.values[0] | row0.values[1] | ...
 *   | row1  | row1.values[0] | ...            | ...
 *
 * Header row is row 1 in the sheet (1-indexed). The first column carries
 * category labels; subsequent columns carry series values.
 */
export const buildEmbeddedXlsx = (
  seriesNames: ReadonlyArray<string>,
  rows: ReadonlyArray<DataRow>,
): Uint8Array => {
  // ----- /xl/worksheets/sheet1.xml --------------------------------------
  const sheetXmlParts: string[] = [];
  sheetXmlParts.push(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<sheetData>',
  );

  // Header row.
  sheetXmlParts.push('<row r="1">');
  // A1 = blank.
  sheetXmlParts.push(`<c r="${cellRef(0, 0)}" t="inlineStr"><is><t></t></is></c>`);
  for (let i = 0; i < seriesNames.length; i++) {
    sheetXmlParts.push(
      `<c r="${cellRef(0, i + 1)}" t="inlineStr"><is><t>${xmlEscape(seriesNames[i] ?? '')}</t></is></c>`,
    );
  }
  sheetXmlParts.push('</row>');

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    const rowIdx = r + 1; // header sits at index 0.
    sheetXmlParts.push(`<row r="${rowIdx + 1}">`);
    sheetXmlParts.push(
      `<c r="${cellRef(rowIdx, 0)}" t="inlineStr"><is><t>${xmlEscape(row.label)}</t></is></c>`,
    );
    for (let i = 0; i < row.values.length; i++) {
      const v = row.values[i];
      if (v === null || v === undefined) continue; // omit empty cell.
      sheetXmlParts.push(`<c r="${cellRef(rowIdx, i + 1)}"><v>${v}</v></c>`);
    }
    sheetXmlParts.push('</row>');
  }

  sheetXmlParts.push('</sheetData></worksheet>');
  const sheetXml = sheetXmlParts.join('');

  // ----- /xl/workbook.xml -----------------------------------------------
  const workbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>';

  // ----- /xl/_rels/workbook.xml.rels ------------------------------------
  const workbookRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" ' +
    'Target="worksheets/sheet1.xml"/></Relationships>';

  // ----- /_rels/.rels ---------------------------------------------------
  const rootRelsXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" ' +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
    'Target="xl/workbook.xml"/></Relationships>';

  // ----- /[Content_Types].xml ------------------------------------------
  const contentTypesXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ' +
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>';

  // Order matters — match what Excel itself emits.
  return writeZip([
    { name: '[Content_Types].xml', data: encode(contentTypesXml) },
    { name: '_rels/.rels', data: encode(rootRelsXml) },
    { name: 'xl/workbook.xml', data: encode(workbookXml) },
    { name: 'xl/_rels/workbook.xml.rels', data: encode(workbookRelsXml) },
    { name: 'xl/worksheets/sheet1.xml', data: encode(sheetXml) },
  ]);
};

/**
 * Returns the spreadsheet-cell reference span used by chart `<c:f>`
 * formulas. Example: `cellRange('Sheet1', 1, 0, 4)` →
 * `"Sheet1!$A$2:$A$5"` (column A, rows 2–5 inclusive, 0-indexed input).
 */
export const cellRange = (sheet: string, startRow: number, col: number, count: number): string => {
  const colA = colLetter(col);
  return `${sheet}!$${colA}$${startRow + 1}:$${colA}$${startRow + count}`;
};

/** Single-cell reference, e.g. `cellAddr('Sheet1', 0, 1)` → `"Sheet1!$B$1"`. */
export const cellAddr = (sheet: string, row: number, col: number): string =>
  `${sheet}!$${colLetter(col)}$${row + 1}`;
