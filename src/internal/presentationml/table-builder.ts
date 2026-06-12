// Builds a table inside a graphic frame.
//
// PPTX wraps tables in `<p:graphicFrame>` whose `<a:graphicData>` carries
// the `<a:tbl>` payload. ECMA-376 Part 1 §19.3.1.21 (CT_GraphicalObjectFrame)
// + §21.1.3.13 (CT_Table). At this phase we expose plain cells with
// default styling; cell-level fill / alignment / span lands when the next
// feature batch needs them.
//
// Column widths and row heights are in EMU. When the caller omits widths,
// we divide the total width equally; same for heights.

import { type XmlElement, NS, attr, elem, qname, text as textNode } from '../xml/index.ts';

const TABLE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';

const NAME_GRAPHIC_FRAME = qname('p', 'graphicFrame', NS.pml);
const NAME_NV_GRAPHIC_FRAME_PR = qname('p', 'nvGraphicFramePr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_GRAPHIC_FRAME_PR = qname('p', 'cNvGraphicFramePr', NS.pml);
const NAME_GRAPHIC_FRAME_LOCKS = qname('a', 'graphicFrameLocks', NS.dml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_P_XFRM = qname('p', 'xfrm', NS.pml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_GRAPHIC = qname('a', 'graphic', NS.dml);
const NAME_GRAPHIC_DATA = qname('a', 'graphicData', NS.dml);
const NAME_TBL = qname('a', 'tbl', NS.dml);
const NAME_TBL_PR = qname('a', 'tblPr', NS.dml);
const NAME_TBL_GRID = qname('a', 'tblGrid', NS.dml);
const NAME_GRID_COL = qname('a', 'gridCol', NS.dml);
const NAME_TR = qname('a', 'tr', NS.dml);
const NAME_TC = qname('a', 'tc', NS.dml);
const NAME_TC_PR = qname('a', 'tcPr', NS.dml);
const NAME_TX_BODY = qname('a', 'txBody', NS.dml);
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_NO_GRP = qname('', 'noGrp', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_URI = qname('', 'uri', '');
const ATTR_W = qname('', 'w', '');
const ATTR_H = qname('', 'h', '');
const ATTR_FIRST_ROW = qname('', 'firstRow', '');
const ATTR_BAND_ROW = qname('', 'bandRow', '');
const ATTR_LANG = qname('', 'lang', '');
const ATTR_XML_SPACE = qname('xml', 'space', NS.xml);

export interface TableOptions {
  id: number;
  name?: string;
  x: number;
  y: number;
  w: number;
  h: number;
  /** Row-major cell contents. Each row must have the same length. */
  rows: ReadonlyArray<ReadonlyArray<string>>;
  /** Column widths in EMU. Defaults to equal distribution of `w`. */
  colWidths?: ReadonlyArray<number>;
  /** Row heights in EMU. Defaults to equal distribution of `h`. */
  rowHeights?: ReadonlyArray<number>;
  /** Emit `firstRow="1"` so the theme styles the header row. Default true. */
  firstRow?: boolean;
  /** Emit `bandRow="1"` so alternating rows get banded shading. Default true. */
  bandRow?: boolean;
}

const buildTextCellBody = (value: string): XmlElement => {
  const needsPreserve =
    value.length > 0 && (value.startsWith(' ') || value.endsWith(' ') || /[\t\n]/.test(value));
  const t = elem(NAME_T, {
    attrs: needsPreserve ? [attr(ATTR_XML_SPACE, 'preserve')] : [],
    children: value.length > 0 ? [textNode(value)] : [],
  });
  const r = elem(NAME_R, {
    children: [elem(NAME_RPR, { attrs: [attr(ATTR_LANG, 'en-US')] }), t],
  });
  const p = elem(NAME_P, { children: [r] });
  return elem(NAME_TX_BODY, { children: [elem(NAME_BODY_PR), elem(NAME_LST_STYLE), p] });
};

/** @internal — used by row-mutation paths in the public API. */
export const buildTableCell = (value: string): XmlElement => {
  return elem(NAME_TC, { children: [buildTextCellBody(value), elem(NAME_TC_PR)] });
};

/** @internal — used by row-mutation paths in the public API. */
export const buildTableRow = (cells: ReadonlyArray<string>, h: number): XmlElement =>
  elem(NAME_TR, {
    attrs: [attr(ATTR_H, String(Math.round(h)))],
    children: cells.map(buildTableCell),
  });

const equalShares = (total: number, n: number): number[] => {
  if (n <= 0) return [];
  // Distribute integer EMU values so they sum exactly to `total` (no
  // 1-EMU drift). The last cell absorbs any rounding remainder.
  const base = Math.floor(total / n);
  const shares = Array.from({ length: n }, () => base);
  shares[n - 1] = base + (total - base * n);
  return shares;
};

export const buildTable = (opts: TableOptions): XmlElement => {
  const rows = opts.rows;
  // Empty rows (or a row with no cells) would emit an `<a:tbl>` with no
  // grid — XML PowerPoint rejects with a repair dialog. Fail loudly at the
  // authoring boundary instead. Errors name the public `addSlideTable`
  // entry point (this builder's sole caller) so the message is actionable.
  if (rows.length === 0) throw new Error('addSlideTable: at least one row is required');
  const colCount = rows[0]?.length ?? 0;
  if (colCount === 0) throw new Error('addSlideTable: at least one column is required');
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r) throw new Error(`addSlideTable: row ${i} is missing`);
    if (r.length !== colCount) {
      throw new Error(
        `addSlideTable: row ${i} has ${r.length} cells; expected ${colCount} to match row 0`,
      );
    }
  }

  const colWidths = opts.colWidths ?? equalShares(opts.w, colCount);
  if (colWidths.length !== colCount) {
    throw new Error(
      `addSlideTable: colWidths has ${colWidths.length} entries; expected ${colCount}`,
    );
  }
  const rowHeights = opts.rowHeights ?? equalShares(opts.h, rows.length);
  if (rowHeights.length !== rows.length) {
    throw new Error(
      `addSlideTable: rowHeights has ${rowHeights.length} entries; expected ${rows.length}`,
    );
  }

  const name = opts.name ?? `Table ${opts.id}`;

  // Non-visual graphic frame properties.
  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const cNvGraphicFramePr = elem(NAME_C_NV_GRAPHIC_FRAME_PR, {
    children: [elem(NAME_GRAPHIC_FRAME_LOCKS, { attrs: [attr(ATTR_NO_GRP, '1')] })],
  });
  const nvPr = elem(NAME_NV_PR);
  const nvGraphicFramePr = elem(NAME_NV_GRAPHIC_FRAME_PR, {
    children: [cNvPr, cNvGraphicFramePr, nvPr],
  });

  // Geometry.
  const off = elem(NAME_OFF, {
    attrs: [attr(ATTR_X, String(opts.x)), attr(ATTR_Y, String(opts.y))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(ATTR_CX, String(opts.w)), attr(ATTR_CY, String(opts.h))],
  });
  const xfrm = elem(NAME_P_XFRM, { children: [off, ext] });

  // Table payload.
  const tblPr = elem(NAME_TBL_PR, {
    attrs: [
      attr(ATTR_FIRST_ROW, (opts.firstRow ?? true) ? '1' : '0'),
      attr(ATTR_BAND_ROW, (opts.bandRow ?? true) ? '1' : '0'),
    ],
  });
  const tblGrid = elem(NAME_TBL_GRID, {
    children: colWidths.map((w) =>
      elem(NAME_GRID_COL, { attrs: [attr(ATTR_W, String(Math.round(w)))] }),
    ),
  });
  const tableRows = rows.map((row, i) => buildTableRow(row, rowHeights[i] ?? 0));
  const tbl = elem(NAME_TBL, { children: [tblPr, tblGrid, ...tableRows] });

  const graphicData = elem(NAME_GRAPHIC_DATA, {
    attrs: [attr(ATTR_URI, TABLE_URI)],
    children: [tbl],
  });
  const graphic = elem(NAME_GRAPHIC, { children: [graphicData] });

  return elem(NAME_GRAPHIC_FRAME, { children: [nvGraphicFramePr, xfrm, graphic] });
};
