// Builds a generic preset shape (`<p:sp>` with `<a:prstGeom prst="...">`).
//
// Different from `text-box-builder` in two ways:
//   - No `txBox="1"` on `<p:cNvSpPr>`; PowerPoint treats this as a regular
//     shape with optional text rather than a free-form text frame.
//   - The geometry preset is caller-chosen from the full ECMA-376
//     `ST_ShapeType` catalog (rect, ellipse, triangle, rightArrow, star5,
//     leftRightArrow, cloud, ...).
//
// Like `text-box-builder`, multi-line text splits on `\n` into paragraphs
// so each `<a:t>` holds a single line (strict schema friendly).

import { emuCoordinate, emuExtent } from '../bounds.ts';
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
const NAME_BODY_PR = qname('a', 'bodyPr', NS.dml);
const NAME_LST_STYLE = qname('a', 'lstStyle', NS.dml);
const NAME_P = qname('a', 'p', NS.dml);
const NAME_R = qname('a', 'r', NS.dml);
const NAME_RPR = qname('a', 'rPr', NS.dml);
const NAME_T = qname('a', 't', NS.dml);
const ATTR_ID = qname('', 'id', '');
const ATTR_NAME = qname('', 'name', '');
const ATTR_X = qname('', 'x', '');
const ATTR_Y = qname('', 'y', '');
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_PRST = qname('', 'prst', '');
const ATTR_WRAP = qname('', 'wrap', '');
const ATTR_RTL_COL = qname('', 'rtlCol', '');
const ATTR_LANG = qname('', 'lang', '');
const ATTR_ANCHOR = qname('', 'anchor', '');
const ATTR_ANCHOR_CTR = qname('', 'anchorCtr', '');

/**
 * Common ECMA-376 `ST_ShapeType` preset tokens. The schema accepts ~180
 * presets; the union below covers the ones authoring scripts reach for
 * most often. Pass any string to `addShape({ preset })` to use a token
 * not enumerated here.
 */
export type PresetShape =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'rtTriangle'
  | 'diamond'
  | 'parallelogram'
  | 'trapezoid'
  | 'pentagon'
  | 'hexagon'
  | 'heptagon'
  | 'octagon'
  | 'decagon'
  | 'star4'
  | 'star5'
  | 'star6'
  | 'star7'
  | 'star8'
  | 'star10'
  | 'star12'
  | 'star16'
  | 'star24'
  | 'star32'
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'leftRightArrow'
  | 'upDownArrow'
  | 'bentArrow'
  | 'curvedRightArrow'
  | 'curvedLeftArrow'
  | 'curvedUpArrow'
  | 'curvedDownArrow'
  | 'cloud'
  | 'heart'
  | 'lightningBolt'
  | 'sun'
  | 'moon'
  | 'bracketPair'
  | 'bracePair'
  | 'leftBracket'
  | 'rightBracket'
  | 'leftBrace'
  | 'rightBrace'
  | 'cube'
  | 'can'
  | 'donut'
  | 'noSmoking'
  | 'plus'
  // ECMA-376 ST_ShapeType spells the math operators `math*`; the bare
  // `minus`/`mult`/`div`/`equal`/`notEqual` are not in the enum and emitted a
  // `prstGeom` PowerPoint silently dropped (the shape vanished on open).
  | 'mathPlus'
  | 'mathMinus'
  | 'mathMultiply'
  | 'mathDivide'
  | 'mathEqual'
  | 'mathNotEqual';

export interface ShapeOptions {
  id: number;
  name?: string;
  preset: PresetShape | string;
  x: number;
  y: number;
  w: number;
  h: number;
  /**
   * Optional text content for the shape. Newlines split into paragraphs;
   * to add formatting / bullets afterwards, call `setTextFormat` /
   * `setBullets` on the returned shape.
   */
  text?: string;
  /**
   * Vertical anchor of any text body (`ST_TextAnchoringType`: `t` top,
   * `ctr` middle, `b` bottom). Defaults to `'ctr'` so preset shapes render
   * with centered text, which is what PowerPoint does when you "insert shape
   * → type text". For horizontal alignment, set the paragraph alignment via
   * `setShapeAlignment` / `setParagraphAlignment` after creating the shape.
   */
  textAnchor?: 'ctr' | 't' | 'b';
}

const buildTextBody = (
  text: string | undefined,
  anchor: ShapeOptions['textAnchor'],
): XmlElement | null => {
  if (text === undefined) return null;
  const bodyPrAttrs = [attr(ATTR_WRAP, 'square'), attr(ATTR_RTL_COL, '0')];
  if (anchor === 'ctr' || anchor === 't' || anchor === 'b') {
    bodyPrAttrs.push(attr(ATTR_ANCHOR, anchor));
  }
  if (anchor === 'ctr' || anchor === undefined) {
    bodyPrAttrs.push(attr(ATTR_ANCHOR_CTR, '1'));
  }
  const bodyPr = elem(NAME_BODY_PR, { attrs: bodyPrAttrs });
  const paragraphs = text.split('\n').map((line) => {
    const t = elem(NAME_T, { children: line.length > 0 ? [textNode(line)] : [] });
    const r = elem(NAME_R, {
      children: [elem(NAME_RPR, { attrs: [attr(ATTR_LANG, 'en-US')] }), t],
    });
    return elem(NAME_P, { children: [r] });
  });
  return elem(NAME_TX_BODY, {
    children: [bodyPr, elem(NAME_LST_STYLE), ...paragraphs],
  });
};

export const buildShape = (opts: ShapeOptions): XmlElement => {
  const name = opts.name ?? `${opts.preset} ${opts.id}`;
  const anchor = opts.textAnchor ?? 'ctr';

  const cNvPr = elem(NAME_C_NV_PR, {
    attrs: [attr(ATTR_ID, String(opts.id)), attr(ATTR_NAME, name)],
  });
  const cNvSpPr = elem(NAME_C_NV_SP_PR);
  const nvSpPr = elem(NAME_NV_SP_PR, {
    children: [cNvPr, cNvSpPr, elem(NAME_NV_PR)],
  });

  // EMU coordinates are integers (ST_Coordinate / xsd:long). Round on the way
  // out so a fractional value from EMU arithmetic (e.g. an `as Emu` cast on a
  // computed fit/translate) can't reach the XML and trip PowerPoint's repair.
  const off = elem(NAME_OFF, {
    attrs: [
      attr(ATTR_X, String(emuCoordinate(opts.x, 'addSlideShape: x'))),
      attr(ATTR_Y, String(emuCoordinate(opts.y, 'addSlideShape: y'))),
    ],
  });
  const ext = elem(NAME_EXT, {
    attrs: [
      attr(ATTR_CX, String(emuExtent(opts.w, 'addSlideShape: w'))),
      attr(ATTR_CY, String(emuExtent(opts.h, 'addSlideShape: h'))),
    ],
  });
  const xfrm = elem(NAME_A_XFRM, { children: [off, ext] });
  const prstGeom = elem(NAME_PRST_GEOM, {
    attrs: [attr(ATTR_PRST, opts.preset)],
    children: [elem(NAME_AV_LST)],
  });
  const spPr = elem(NAME_SP_PR, { children: [xfrm, prstGeom] });

  const children: XmlElement[] = [nvSpPr, spPr];
  const txBody = buildTextBody(opts.text, anchor);
  if (txBody) children.push(txBody);

  return elem(NAME_SP, { children });
};
