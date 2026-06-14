// Builds the XML element for a free-form text box (`<p:sp>` with
// `txBox="1"` on `cNvSpPr` and a `prstGeom prst="rect"` shape).
//
// This is the simplest authoring shape: a rectangle with no fill or
// outline, containing a single paragraph of text. Geometry comes from
// the caller in EMU. Text properties beyond a single run are deferred to
// future iterations.

import { type XmlElement, NS, attr, elem, qname, text as textNode } from '../xml/index.ts';

const NAME_SP = qname('p', 'sp', NS.pml);
const NAME_NV_SP_PR = qname('p', 'nvSpPr', NS.pml);
const NAME_C_NV_PR = qname('p', 'cNvPr', NS.pml);
const NAME_C_NV_SP_PR = qname('p', 'cNvSpPr', NS.pml);
const NAME_NV_PR = qname('p', 'nvPr', NS.pml);
const NAME_SP_PR = qname('p', 'spPr', NS.pml);
const NAME_TX_BODY = qname('p', 'txBody', NS.pml);
const NAME_A_XFRM = qname('a', 'xfrm', NS.dml);
const NAME_OFF = qname('a', 'off', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const NAME_PRST_GEOM = qname('a', 'prstGeom', NS.dml);
const NAME_AV_LST = qname('a', 'avLst', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_TX_BOX = qname('', 'txBox', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_PRST = qname('', 'prst', '');
const ATTR_WRAP = qname('', 'wrap', '');
const ATTR_RTL_COL = qname('', 'rtlCol', '');
const ATTR_LANG = qname('', 'lang', '');
const ATTR_XML_SPACE = qname('xml', 'space', NS.xml);

export interface TextBoxOptions {
  /** Shape numeric id, unique within the slide. */
  id: number;
  /** Shape user-visible name. Defaults to `TextBox <id>`. */
  name?: string;
  /** Position (left edge), in EMU. */
  x: number;
  /** Position (top edge), in EMU. */
  y: number;
  /** Width, in EMU. */
  w: number;
  /** Height, in EMU. */
  h: number;
  /** The text content. Newlines are NOT split into paragraphs here; for
   * multi-paragraph text, set the body and then call `setText` on the
   * returned shape. */
  text: string;
}

/**
 * Returns a `<p:sp>` element representing a text-box shape positioned at
 * `(x, y)` with extent `(w, h)`. The shape carries a `txBox="1"` marker
 * and a `prstGeom prst="rect"` so PowerPoint renders it as a plain text
 * frame without fill or outline.
 */
export const buildTextBox = (opts: TextBoxOptions): XmlElement => {
  const name = opts.name ?? `TextBox ${opts.id}`;

  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const cNvSpPr = elem(NAME_C_NV_SP_PR, {
    attrs: [attr(ATTR_TX_BOX, '1')],
  });
  const nvPr = elem(NAME_NV_PR);
  const nvSpPr = elem(NAME_NV_SP_PR, { children: [cNvPr, cNvSpPr, nvPr] });

  // Round to whole EMU; fractional ST_Coordinate values corrupt the file.
  const off = elem(NAME_OFF, {
    attrs: [attr(ATTR_X, String(Math.round(opts.x))), attr(ATTR_Y, String(Math.round(opts.y)))],
  });
  const ext = elem(NAME_EXT, {
    attrs: [attr(ATTR_CX, String(Math.round(opts.w))), attr(ATTR_CY, String(Math.round(opts.h)))],
  });
  const xfrm = elem(NAME_A_XFRM, { children: [off, ext] });
  const prstGeom = elem(NAME_PRST_GEOM, {
    attrs: [attr(ATTR_PRST, 'rect')],
    children: [elem(NAME_AV_LST)],
  });
  const spPr = elem(NAME_SP_PR, {
    children: [xfrm, prstGeom, elem(NAME_NO_FILL)],
  });

  // Build one `<a:p>` per line of `opts.text`. Each `<a:t>` therefore
  // contains at most one line, so we never need `xml:space="preserve"`
  // (the strict ECMA schema disallows that attribute on `<a:t>`, even
  // though PowerPoint sometimes emits it). Leading/trailing-space
  // preservation is handled by the bodyPr, not by xml:space.
  const lines = opts.text.split('\n');
  const paragraphs: XmlElement[] = lines.map((line) => {
    const t = elem(NAME_T, {
      children: line.length > 0 ? [textNode(line)] : [],
    });
    const r = elem(NAME_R, {
      children: [elem(NAME_RPR, { attrs: [attr(ATTR_LANG, 'en-US')] }), t],
    });
    return elem(NAME_P, { children: [r] });
  });
  const bodyPr = elem(NAME_BODY_PR, {
    attrs: [attr(ATTR_WRAP, 'square'), attr(ATTR_RTL_COL, '0')],
  });
  const txBody = elem(NAME_TX_BODY, {
    children: [bodyPr, elem(NAME_LST_STYLE), ...paragraphs],
  });
  void ATTR_XML_SPACE;

  return elem(NAME_SP, { children: [nvSpPr, spPr, txBody] });
};
