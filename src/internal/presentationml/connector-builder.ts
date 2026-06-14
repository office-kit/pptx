// Builds a straight-line connector shape (`<p:cxnSp>` with
// `<a:prstGeom prst="line">`).
//
// PPTX represents a connector as a shape whose bounding box has its
// upper-left corner at `(min(x1,x2), min(y1,y2))` and extent of
// `(|x2-x1|, |y2-y1|)`. Direction (which end is "from" and which is
// "to") is captured via `flipH` / `flipV` on the xfrm.

import { type XmlElement, NS, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from '../drawingml/color.ts';

const NAME_CXN_SP = qname('p', 'cxnSp', NS.pml);
const NAME_NV_CXN_SP_PR = qname('p', 'nvCxnSpPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_CXN_SP_PR = qname('p', 'cNvCxnSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_A_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_PRST_GEOM = qname('a', 'prstGeom', NS.dml);
const NAME_AV_LST = qname('a', 'avLst', NS.dml);
const NAME_LN = qname('a', 'ln', NS.dml);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_PRST = qname('', 'prst', '');
const ATTR_W = qname('', 'w', '');
const ATTR_FLIP_H = qname('', 'flipH', '');
const ATTR_FLIP_V = qname('', 'flipV', '');

export interface ConnectorOptions {
  id: number;
  name?: string;
  /** Start point in EMU. */
  from: { x: number; y: number };
  /** End point in EMU. */
  to: { x: number; y: number };
  /**
   * Line color. Same accepted forms as `setFill` — `#RRGGBB`, scheme
   * tokens, etc. When omitted PowerPoint applies the layout's default
   * line color.
   */
  color?: string;
  /** Line width in EMU. PowerPoint's hairline default is `9525` (0.75pt). */
  widthEmu?: number;
}

/** Returns a `<p:cxnSp>` straight-line connector. */
export const buildConnector = (opts: ConnectorOptions): XmlElement => {
  const name = opts.name ?? `Straight Connector ${opts.id}`;

  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const cNvCxnSpPr = elem(NAME_C_NV_CXN_SP_PR);
  const nvCxnSpPr = elem(NAME_NV_CXN_SP_PR, {
    children: [cNvPr, cNvCxnSpPr, elem(NAME_NV_PR)],
  });

  const x = Math.min(opts.from.x, opts.to.x);
  const y = Math.min(opts.from.y, opts.to.y);
  const cx = Math.abs(opts.to.x - opts.from.x);
  const cy = Math.abs(opts.to.y - opts.from.y);
  const flipH = opts.from.x > opts.to.x;
  const flipV = opts.from.y > opts.to.y;

  const xfrmAttrs = [];
  if (flipH) xfrmAttrs.push(attr(ATTR_FLIP_H, '1'));
  if (flipV) xfrmAttrs.push(attr(ATTR_FLIP_V, '1'));
  const xfrm = elem(NAME_A_XFRM, {
    attrs: xfrmAttrs,
    children: [
      // Round to whole EMU; fractional ST_Coordinate values corrupt the file.
      elem(NAME_OFF, {
        attrs: [attr(ATTR_X, String(Math.round(x))), attr(ATTR_Y, String(Math.round(y)))],
      }),
      elem(NAME_EXT, {
        attrs: [attr(ATTR_CX, String(Math.round(cx))), attr(ATTR_CY, String(Math.round(cy)))],
      }),
    ],
  });
  const prstGeom = elem(NAME_PRST_GEOM, {
    attrs: [attr(ATTR_PRST, 'line')],
    children: [elem(NAME_AV_LST)],
  });

  const spPrChildren: XmlElement[] = [xfrm, prstGeom];
  if (opts.color !== undefined || opts.widthEmu !== undefined) {
    const ln = elem(NAME_LN, {
      attrs: opts.widthEmu !== undefined ? [attr(ATTR_W, String(Math.round(opts.widthEmu)))] : [],
      children:
        opts.color !== undefined
          ? [elem(NAME_SOLID_FILL, { children: [buildColorElement(opts.color)] })]
          : [],
    });
    spPrChildren.push(ln);
  }
  const spPr = elem(NAME_SP_PR, { children: spPrChildren });

  return elem(NAME_CXN_SP, { children: [nvCxnSpPr, spPr] });
};
