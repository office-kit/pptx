// Table cell access.

import { resolveChartPartName } from './charts.ts';
import {
  applyAlignmentToAllParagraphs,
  applyFormatToAllRuns,
  clearFill as clearFillImpl,
  setSolidFill,
  setTextBody,
  type TextFormat,
  type ParagraphAlignment,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import { buildTableCell, buildTableRow } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  allChildElements,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  CELL_COL,
  CELL_ELEMENT,
  CELL_ROW,
  CELL_TABLE,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  type SlideShapeData,
  type TableCellData,
} from '../_internal-symbols.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
import { getPresentationTheme } from './package.ts';
import { resolveDrawingColor } from './shapes.ts';

// ---------------------------------------------------------------------------
// Table cell access.
//
// `addSlideTable` builds the cell tree under `<a:graphic>/<a:graphicData>/
// <a:tbl>`. These helpers let callers reach individual `<a:tc>` cells to
// re-fill text, paint backgrounds, or align contents without rebuilding
// the table.

export type { TableCellData };

const NAME_A_GRAPHIC_TBL = qname('a', 'graphic', NS.dml);
const NAME_A_GRAPHIC_DATA_TBL = qname('a', 'graphicData', NS.dml);
const NAME_A_TBL = qname('a', 'tbl', NS.dml);
const NAME_A_TC_PR = qname('a', 'tcPr', NS.dml);
const NAME_A_TX_BODY_TBL = qname('a', 'txBody', NS.dml);

const findTblElement = (shape: SlideShapeData): XmlElement | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'graphicFrame') return null;
  const graphic = firstChildElement(shape[SHAPE_ELEMENT], NAME_A_GRAPHIC_TBL);
  if (!graphic) return null;
  const graphicData = firstChildElement(graphic, NAME_A_GRAPHIC_DATA_TBL);
  if (!graphicData) return null;
  return firstChildElement(graphicData, NAME_A_TBL);
};

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping `<a:tbl>` —
 * i.e. a table. Sharper than `getShapeKind(shape) === 'graphicFrame'`,
 * which also matches charts and SmartArt frames.
 */
export const isTableShape = (shape: SlideShapeData): boolean => findTblElement(shape) !== null;

/**
 * `true` when the shape is a `<p:graphicFrame>` wrapping a chart
 * reference (`<c:chart>`). Charts, tables, and SmartArt all share
 * the graphic-frame kind; this predicate filters down to charts only.
 */
export const isChartShape = (shape: SlideShapeData): boolean =>
  resolveChartPartName(shape[SHAPE_SLIDE], shape) !== null;

const NAME_A_GRID_COL = qname('a', 'gridCol', NS.dml);
const ATTR_W_TBL = qname('', 'w', '');
const ATTR_H_TBL = qname('', 'h', '');

const tableRows = (tbl: XmlElement): XmlElement[] =>
  tbl.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tr',
  );

const rowCells = (tr: XmlElement): XmlElement[] =>
  tr.children.filter(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tc',
  );

const buildCellHandle = (
  table: SlideShapeData,
  tc: XmlElement,
  row: number,
  col: number,
): TableCellData => ({
  [CELL_TABLE]: table,
  [CELL_ELEMENT]: tc,
  [CELL_ROW]: row,
  [CELL_COL]: col,
});

/**
 * Returns a 2D array of cell handles for the given table shape, in
 * row-major order. Throws if the shape isn't a table graphic frame.
 */
export const getTableCells = (
  table: SlideShapeData,
): ReadonlyArray<ReadonlyArray<TableCellData>> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableCells: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  return rows.map((tr, rowIdx) =>
    rowCells(tr).map((tc, colIdx) => buildCellHandle(table, tc, rowIdx, colIdx)),
  );
};

/**
 * Reads the table-style GUID from `<a:tbl><a:tblPr><a:tableStyleId>`.
 * PowerPoint uses GUIDs to reference built-in table styles
 * (`{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` = "Medium Style 2 -
 * Accent 1", etc.) and theme-local styles. Returns `null` when the
 * table doesn't reference one (uses the slide's default style).
 */
export const getTableStyleId = (table: SlideShapeData): string | null => {
  const tbl = findTblElement(table);
  if (!tbl) return null;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return null;
  const idEl = firstChildElement(tblPr, qname('a', 'tableStyleId', NS.dml));
  if (!idEl) return null;
  let acc = '';
  for (const c of idEl.children) {
    if (c.kind === 'text' || c.kind === 'cdata') acc += c.data;
  }
  return acc.trim() || null;
};

/**
 * Reads the boolean style flags off `<a:tblPr>` — which header /
 * footer rows + columns are banded or emphasised. Mirrors the
 * `firstRow` / `bandRow` flags exposed by `addSlideTable`.
 *
 * Per ECMA-376 §17.18.95 / §21.1.3.15:
 *
 *   - `firstRow` — header row is styled differently.
 *   - `lastRow` — total row is styled differently.
 *   - `firstCol` — first column is styled differently.
 *   - `lastCol` — last column is styled differently.
 *   - `bandRow` — alternating row shading.
 *   - `bandCol` — alternating column shading.
 *
 * Renderers use these to switch on the corresponding table style
 * variant. Returns all-`false` for tables that don't author `<a:tblPr>`.
 */
export const getTableStyleFlags = (
  table: SlideShapeData,
): {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  bandRow: boolean;
  bandCol: boolean;
} => {
  const empty = {
    firstRow: false,
    lastRow: false,
    firstCol: false,
    lastCol: false,
    bandRow: false,
    bandCol: false,
  };
  const tbl = findTblElement(table);
  if (!tbl) return empty;
  const tblPr = firstChildElement(tbl, qname('a', 'tblPr', NS.dml));
  if (!tblPr) return empty;
  const readBool = (attr: string): boolean => {
    const v = getAttrValue(tblPr, qname('', attr, ''));
    return v === '1' || v === 'true';
  };
  return {
    firstRow: readBool('firstRow'),
    lastRow: readBool('lastRow'),
    firstCol: readBool('firstCol'),
    lastCol: readBool('lastCol'),
    bandRow: readBool('bandRow'),
    bandCol: readBool('bandCol'),
  };
};

/**
 * Returns the table's row + column counts. Throws when the shape
 * isn't a table graphic frame.
 */
export const getTableDimensions = (
  table: SlideShapeData,
): { readonly rows: number; readonly cols: number } => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableDimensions: shape is not a table graphic frame');
  const rows = tableRows(tbl);
  const cols = rows[0] !== undefined ? rowCells(rows[0]).length : 0;
  return { rows: rows.length, cols };
};

/**
 * Returns each column's width in EMU, in left-to-right order, as
 * declared on `<a:tblGrid>/<a:gridCol w="...">`. Missing or
 * unparseable widths default to 0. Throws when the shape isn't a
 * table graphic frame.
 */
export const getTableColumnWidths = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableColumnWidths: shape is not a table graphic frame');
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return [];
  const out: Emu[] = [];
  for (const col of allChildElements(grid, NAME_A_GRID_COL)) {
    const v = getAttrValue(col, ATTR_W_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Returns the table's nominal `(width, height)` derived from
 * summing the `<a:gridCol w>` and `<a:tr h>` attributes (both in
 * EMU). Useful for layout pipelines that want to know how big a
 * table really is without dereferencing the shape's `<a:xfrm>`.
 *
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableSize = (
  table: SlideShapeData,
): { readonly width: Emu; readonly height: Emu } => {
  const widths = getTableColumnWidths(table);
  const heights = getTableRowHeights(table);
  const width = widths.reduce((sum, w) => sum + w, 0) as Emu;
  const height = heights.reduce((sum, h) => sum + h, 0) as Emu;
  return { width, height };
};

/**
 * Returns each row's height in EMU, in top-to-bottom order, from
 * `<a:tr h="...">`. Missing or unparseable heights default to 0.
 * Throws when the shape isn't a table graphic frame.
 */
export const getTableRowHeights = (table: SlideShapeData): ReadonlyArray<Emu> => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('getTableRowHeights: shape is not a table graphic frame');
  const out: Emu[] = [];
  for (const tr of tableRows(tbl)) {
    const v = getAttrValue(tr, ATTR_H_TBL);
    const n = v !== null ? Number.parseInt(v, 10) : 0;
    out.push((Number.isFinite(n) ? n : 0) as Emu);
  }
  return out;
};

/**
 * Sets a single column's width on the table grid. Throws on
 * out-of-range column indices or non-table shapes. The total table
 * width is not auto-adjusted — callers are responsible for keeping
 * the sum consistent with the table's `<a:xfrm>` extent if PowerPoint
 * is to render the table without clipping.
 */
export const setTableColumnWidth = (table: SlideShapeData, col: number, width: Emu): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table has no <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const target = cols[col];
  if (!target) throw new RangeError(`table column ${col} out of range (have ${cols.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'w'),
  );
  target.attrs.push(attr(ATTR_W_TBL, String(Math.round(width))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Sets a single row's height. Throws on out-of-range row indices or
 * non-table shapes. As with `setTableColumnWidth`, the table's
 * `<a:xfrm>` extent is left to the caller.
 */
export const setTableRowHeight = (table: SlideShapeData, row: number, height: Emu): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  const target = rows[row];
  if (!target) throw new RangeError(`table row ${row} out of range (have ${rows.length})`);
  target.attrs = target.attrs.filter(
    (a) => !(a.name.namespaceURI === '' && a.name.localName === 'h'),
  );
  target.attrs.push(attr(ATTR_H_TBL, String(Math.round(height))));
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Returns the cell at `(row, col)`. Throws on out-of-range coordinates
 * or non-table shapes.
 */
export const getTableCell = (table: SlideShapeData, row: number, col: number): TableCellData => {
  const cells = getTableCells(table);
  const r = cells[row];
  if (!r) throw new RangeError(`table row ${row} out of range (have ${cells.length})`);
  const c = r[col];
  if (!c) throw new RangeError(`table column ${col} out of range in row ${row} (have ${r.length})`);
  return c;
};

const commitTableCell = (cell: TableCellData): void => {
  const shape = cell[CELL_TABLE];
  commitSlideData(shape[SHAPE_SLIDE]);
  refreshSlideData(shape[SHAPE_SLIDE]);
};

const ensureCellTxBody = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let txBody = firstChildElement(tc, NAME_A_TX_BODY_TBL);
  if (txBody === null) {
    txBody = elem(NAME_A_TX_BODY_TBL);
    // bodyPr lstStyle a:p — keep the canonical ordering.
    txBody.children.push(elem(qname('a', 'bodyPr', NS.dml)));
    txBody.children.push(elem(qname('a', 'lstStyle', NS.dml)));
    // Insert before <a:tcPr> per the schema.
    const tcPrIdx = tc.children.findIndex(
      (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'tcPr',
    );
    if (tcPrIdx >= 0) tc.children.splice(tcPrIdx, 0, txBody);
    else tc.children.push(txBody);
  }
  return txBody;
};

const ensureCellTcPr = (cell: TableCellData): XmlElement => {
  const tc = cell[CELL_ELEMENT];
  let tcPr = firstChildElement(tc, NAME_A_TC_PR);
  if (tcPr === null) {
    tcPr = elem(NAME_A_TC_PR);
    // <a:tcPr> is the LAST child of <a:tc>.
    tc.children.push(tcPr);
  }
  return tcPr;
};

/** Replaces a cell's text. `\n` starts a new paragraph. */
export const setTableCellText = (cell: TableCellData, text: string): void => {
  const txBody = ensureCellTxBody(cell);
  setTextBody(txBody, text);
  commitTableCell(cell);
};

/**
 * Reads the cell's merge / span attributes:
 *
 *   - `gridSpan` — number of columns this cell spans (≥1, default 1).
 *   - `rowSpan` — number of rows this cell spans (≥1, default 1).
 *   - `hMerge` — `true` when this cell is the right half of a horizontal
 *     span (it's painted by an earlier cell with `gridSpan > 1`).
 *   - `vMerge` — `true` when this cell is the bottom half of a vertical
 *     span (it's painted by an earlier cell with `rowSpan > 1`).
 *
 * Renderers should skip painting cells where `hMerge` or `vMerge` is
 * true; those cells exist only so the row/column grid stays consistent.
 */
export const getTableCellSpan = (
  cell: TableCellData,
): { gridSpan: number; rowSpan: number; hMerge: boolean; vMerge: boolean } => {
  const el = cell[CELL_ELEMENT];
  const gs = getAttrValue(el, qname('', 'gridSpan', ''));
  const rs = getAttrValue(el, qname('', 'rowSpan', ''));
  const hm = getAttrValue(el, qname('', 'hMerge', ''));
  const vm = getAttrValue(el, qname('', 'vMerge', ''));
  const parseSpan = (v: string | null): number => {
    if (v === null) return 1;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) && n > 0 ? n : 1;
  };
  return {
    gridSpan: parseSpan(gs),
    rowSpan: parseSpan(rs),
    hMerge: hm === '1' || hm === 'true',
    vMerge: vm === '1' || vm === 'true',
  };
};

/**
 * One side of a cell's border, as read from `<a:tcPr><a:ln{L|R|T|B}>`.
 * `widthEmu` is the line width in EMU; `color` is `#RRGGBB` or `null`
 * when the border is inherited / un-styled.
 */
export interface TableCellBorder {
  readonly color: string | null;
  readonly widthEmu: number | null;
  readonly dash: string | null;
}

export interface TableCellBorders {
  readonly left: TableCellBorder | null;
  readonly right: TableCellBorder | null;
  readonly top: TableCellBorder | null;
  readonly bottom: TableCellBorder | null;
  readonly tlToBr: TableCellBorder | null;
  readonly blToTr: TableCellBorder | null;
}

/**
 * Reads the per-side borders on a cell. Returns `null` for sides with
 * no explicit `<a:ln{Side}>` element (those inherit from the table
 * style / theme). All four cardinal sides plus the two diagonals
 * (TL→BR, BL→TR) are surfaced because real templates use them all.
 */
export const getTableCellBorders = (
  pres: PresentationData,
  cell: TableCellData,
): TableCellBorders => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const theme = getPresentationTheme(pres);
  const empty: TableCellBorders = {
    left: null,
    right: null,
    top: null,
    bottom: null,
    tlToBr: null,
    blToTr: null,
  };
  if (!tcPr) return empty;
  const readLn = (local: string): TableCellBorder | null => {
    const ln = firstChildElement(tcPr, qname('a', local, NS.dml));
    if (!ln) return null;
    const w = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = w !== null ? Number.parseInt(w, 10) : null;
    let color: string | null = null;
    const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
    if (solid) {
      for (const c of solid.children) {
        if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
        color = resolveDrawingColor(c, theme);
        break;
      }
    }
    const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
    const dash = prstDash ? getAttrValue(prstDash, qname('', 'val', '')) : null;
    return { color, widthEmu, dash };
  };
  return {
    left: readLn('lnL'),
    right: readLn('lnR'),
    top: readLn('lnT'),
    bottom: readLn('lnB'),
    tlToBr: readLn('lnTlToBr'),
    blToTr: readLn('lnBlToTr'),
  };
};

/**
 * Reads the cell's text direction (`<a:tcPr vert="…"/>`) — same tokens
 * as `getShapeTextDirection`. Returns `null` for the default
 * horizontal direction.
 *
 * Vertical column headers in tables almost always emit `<a:tcPr
 * vert="vert270"/>` or `"eaVert"` so the header label reads bottom-to-
 * top alongside its column.
 */
export const getTableCellTextDirection = (
  cell: TableCellData,
): 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const v = getAttrValue(tcPr, qname('', 'vert', ''));
  if (
    v === 'vert' ||
    v === 'vert270' ||
    v === 'wordArtVert' ||
    v === 'eaVert' ||
    v === 'mongolianVert' ||
    v === 'wordArtVertRtl'
  )
    return v;
  return null;
};

/**
 * Reads the cell's vertical text anchor (`<a:tcPr anchor="t|ctr|b"/>`)
 * — `'top'`, `'center'`, `'bottom'`, or `null` for the default
 * (`ctr` / center per the schema).
 */
export const getTableCellAnchor = (cell: TableCellData): 'top' | 'center' | 'bottom' | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const v = getAttrValue(tcPr, qname('', 'anchor', ''));
  if (v === 't') return 'top';
  if (v === 'ctr') return 'center';
  if (v === 'b') return 'bottom';
  return null;
};

/**
 * Reads the cell's inset margins (`<a:tcPr marL marR marT marB>`) in
 * EMU. Each side is `null` when the cell doesn't author it (renderers
 * should fall back to PowerPoint's defaults — 91440 EMU / 0.1 inch
 * for the horizontal margins, 45720 EMU for the vertical).
 */
export const getTableCellMargins = (
  cell: TableCellData,
): { left: number | null; right: number | null; top: number | null; bottom: number | null } => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  const empty = { left: null, right: null, top: null, bottom: null };
  if (!tcPr) return empty;
  const read = (name: string): number | null => {
    const v = getAttrValue(tcPr, qname('', name, ''));
    if (v === null) return null;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n : null;
  };
  return {
    left: read('marL'),
    right: read('marR'),
    top: read('marT'),
    bottom: read('marB'),
  };
};

/** Reads the cell's plain text (paragraphs joined with `\n`). */
export const getTableCellText = (cell: TableCellData): string => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return '';
  const lines: string[] = [];
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    let line = '';
    for (const r of p.children) {
      if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r')
        continue;
      const tEl = firstChildElement(r, qname('a', 't', NS.dml));
      if (!tEl) continue;
      for (const child of tEl.children) {
        if (child.kind === 'text' || child.kind === 'cdata') line += child.data;
      }
    }
    lines.push(line);
  }
  return lines.join('\n');
};

/** Sets a solid background color on a cell (`<a:tcPr><a:solidFill>`). */
export const setTableCellFill = (cell: TableCellData, color: string): void => {
  const tcPr = ensureCellTcPr(cell);
  setSolidFill(tcPr, color);
  commitTableCell(cell);
};

/** Removes any background fill from a cell. */
export const clearTableCellFill = (cell: TableCellData): void => {
  const tcPr = ensureCellTcPr(cell);
  clearFillImpl(tcPr);
  commitTableCell(cell);
};

/**
 * Reads the cell's solid background color. Returns `#RRGGBB` for
 * `<a:srgbClr>`, `scheme:<token>` for `<a:schemeClr>`, or `null` when
 * the cell has no fill, no `<a:tcPr>`, or the fill is non-solid
 * (gradient / pattern / image).
 */
export const getTableCellFill = (cell: TableCellData): string | null => {
  const tcPr = firstChildElement(cell[CELL_ELEMENT], NAME_A_TC_PR);
  if (!tcPr) return null;
  const solid = firstChildElement(tcPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  const srgb = firstChildElement(solid, qname('a', 'srgbClr', NS.dml));
  if (srgb) {
    const v = getAttrValue(srgb, qname('', 'val', ''));
    if (v) return `#${v.toUpperCase()}`;
  }
  const scheme = firstChildElement(solid, qname('a', 'schemeClr', NS.dml));
  if (scheme) {
    const v = getAttrValue(scheme, qname('', 'val', ''));
    if (v) return `scheme:${v}`;
  }
  return null;
};

/** Applies a TextFormat to every run in the cell's text. */
export const setTableCellTextFormat = (cell: TableCellData, format: TextFormat): void => {
  const txBody = ensureCellTxBody(cell);
  applyFormatToAllRuns(txBody, format);
  commitTableCell(cell);
};

/** Sets horizontal alignment on every paragraph in the cell. */
export const setTableCellAlignment = (cell: TableCellData, align: ParagraphAlignment): void => {
  const txBody = ensureCellTxBody(cell);
  applyAlignmentToAllParagraphs(txBody, align);
  commitTableCell(cell);
};

/** Zero-based (row, col) of the cell. */
export const getTableCellPosition = (cell: TableCellData): { row: number; col: number } => ({
  row: cell[CELL_ROW],
  col: cell[CELL_COL],
});

/**
 * Reads the horizontal alignment from the cell's first paragraph
 * (`l`, `ctr`, `r`, `just`, `dist`, `justLow`, `thaiDist`). Returns
 * `null` when the cell has no `<a:txBody>` or its first paragraph
 * has no explicit `algn` attribute (PowerPoint then defaults to `l`).
 */
export const getTableCellAlignment = (cell: TableCellData): ParagraphAlignment | null => {
  const txBody = firstChildElement(cell[CELL_ELEMENT], NAME_A_TX_BODY_TBL);
  if (!txBody) return null;
  for (const p of txBody.children) {
    if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p')
      continue;
    const pPr = firstChildElement(p, qname('a', 'pPr', NS.dml));
    if (!pPr) return null;
    const v = getAttrValue(pPr, qname('', 'algn', ''));
    return (v as ParagraphAlignment | null) ?? null;
  }
  return null;
};

const requireTbl = (table: SlideShapeData): XmlElement => {
  const tbl = findTblElement(table);
  if (!tbl) throw new Error('table shape is not a table graphic frame');
  return tbl;
};

const tableColumnCount = (tbl: XmlElement): number => {
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) return 0;
  return allChildElements(grid, NAME_A_GRID_COL).length;
};

const rowDefaultHeight = (tbl: XmlElement): number => {
  // Use the average height of existing rows, or 370000 (≈ 0.4in) as a
  // sane default when the table has no rows yet.
  const rows = tableRows(tbl);
  if (rows.length === 0) return 370000;
  let sum = 0;
  let count = 0;
  for (const r of rows) {
    const h = getAttrValue(r, ATTR_H_TBL);
    if (h !== null) {
      const n = Number.parseInt(h, 10);
      if (Number.isFinite(n)) {
        sum += n;
        count++;
      }
    }
  }
  return count > 0 ? Math.round(sum / count) : 370000;
};

/**
 * Inserts a row into the table. `atIndex` is 0-based; `undefined`
 * appends at the end. `cells` supplies cell values; missing cells
 * become blank, extras are dropped. The row's height matches the
 * average of existing rows (or a 0.4in default for empty tables).
 */
export const insertTableRow = (
  table: SlideShapeData,
  atIndex?: number,
  cells: ReadonlyArray<string> = [],
): void => {
  const tbl = requireTbl(table);
  const colCount = tableColumnCount(tbl);
  const padded: string[] = [];
  for (let i = 0; i < colCount; i++) padded.push(cells[i] ?? '');
  const row = buildTableRow(padded, rowDefaultHeight(tbl));

  const rows = tableRows(tbl);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, rows.length)) : rows.length;
  if (insertAt === rows.length) {
    tbl.children.push(row);
  } else {
    const target = rows[insertAt]!;
    const idx = tbl.children.indexOf(target);
    tbl.children.splice(idx, 0, row);
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the row at `atIndex` from the table. Throws on out-of-range. */
export const removeTableRow = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const rows = tableRows(tbl);
  if (atIndex < 0 || atIndex >= rows.length) {
    throw new RangeError(`removeTableRow: index ${atIndex} out of range (have ${rows.length})`);
  }
  const target = rows[atIndex]!;
  tbl.children = tbl.children.filter((c) => c !== target);
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/**
 * Inserts a column into the table. `atIndex` defaults to the end.
 * `widthEmu` defaults to the average of existing column widths (or
 * 914400 = 1in if the table has no columns). Existing rows get a new
 * blank cell at `atIndex`.
 */
export const insertTableColumn = (
  table: SlideShapeData,
  atIndex?: number,
  widthEmu?: number,
): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  const insertAt =
    atIndex !== undefined ? Math.max(0, Math.min(atIndex, cols.length)) : cols.length;

  // Default width: average of existing widths.
  let defaultWidth = widthEmu;
  if (defaultWidth === undefined) {
    let sum = 0;
    let count = 0;
    for (const col of cols) {
      const w = getAttrValue(col, ATTR_W_TBL);
      if (w !== null) {
        const n = Number.parseInt(w, 10);
        if (Number.isFinite(n)) {
          sum += n;
          count++;
        }
      }
    }
    defaultWidth = count > 0 ? Math.round(sum / count) : 914400;
  }
  const newCol = elem(NAME_A_GRID_COL, { attrs: [attr(ATTR_W_TBL, String(defaultWidth))] });
  if (insertAt === cols.length) {
    grid.children.push(newCol);
  } else {
    const target = cols[insertAt]!;
    const idx = grid.children.indexOf(target);
    grid.children.splice(idx, 0, newCol);
  }

  // Add a blank <a:tc> at the same column index in every row.
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    const newCell = buildTableCell('');
    if (insertAt >= tcs.length) {
      tr.children.push(newCell);
    } else {
      const target = tcs[insertAt]!;
      const idx = tr.children.indexOf(target);
      tr.children.splice(idx, 0, newCell);
    }
  }

  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};

/** Removes the column at `atIndex` (and the corresponding cell in every row). */
export const removeTableColumn = (table: SlideShapeData, atIndex: number): void => {
  const tbl = requireTbl(table);
  const grid = firstChildElement(tbl, qname('a', 'tblGrid', NS.dml));
  if (!grid) throw new Error('table is missing <a:tblGrid>');
  const cols = allChildElements(grid, NAME_A_GRID_COL);
  if (atIndex < 0 || atIndex >= cols.length) {
    throw new RangeError(`removeTableColumn: index ${atIndex} out of range (have ${cols.length})`);
  }
  const targetCol = cols[atIndex]!;
  grid.children = grid.children.filter((c) => c !== targetCol);
  for (const tr of tableRows(tbl)) {
    const tcs = rowCells(tr);
    if (atIndex < tcs.length) {
      tr.children = tr.children.filter((c) => c !== tcs[atIndex]);
    }
  }
  commitSlideData(table[SHAPE_SLIDE]);
  refreshSlideData(table[SHAPE_SLIDE]);
};
