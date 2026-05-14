// Outline (line) mutation for shapes.
//
// `<a:ln>` sits inside `<p:spPr>` after the fill choice. ECMA-376 §20.1.2
// surface: width (EMU), cap, dash, fill choice (solid/no/grad), head/tail
// arrow markers. At this phase we expose width + solid color + noFill;
// dashes and arrowheads land when the next feature batch needs them.

import { NS, type XmlElement, attr, elem, qname } from '../xml/index.ts';
import { buildColorElement } from './color.ts';

const NAME_LN = qname('a', 'ln', NS.dml);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);
const ATTR_W = qname('', 'w', '');

const FILL_LOCALS = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill']);

const removeChildrenIn = (host: XmlElement, names: ReadonlySet<string>): void => {
  host.children = host.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && names.has(c.name.localName)),
  );
};

const ensureLn = (spPr: XmlElement): XmlElement => {
  for (const c of spPr.children) {
    if (c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'ln') {
      return c;
    }
  }
  // <a:ln> goes AFTER the fill choice and BEFORE effects / scene3d / sp3d /
  // extLst per the schema. We insert at the index of the first element that
  // belongs after `<a:ln>`; otherwise append.
  const ln = elem(NAME_LN);
  const afterLn = new Set(['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']);
  for (let i = 0; i < spPr.children.length; i++) {
    const c = spPr.children[i];
    if (c?.kind === 'element' && c.name.namespaceURI === NS.dml && afterLn.has(c.name.localName)) {
      spPr.children.splice(i, 0, ln);
      return ln;
    }
  }
  spPr.children.push(ln);
  return ln;
};

export interface StrokeOptions {
  /** Line color. Same accepted forms as `setFill`. */
  color?: string;
  /** Line width in EMU. PowerPoint's default for a hairline is 9525 (0.75pt). */
  widthEmu?: number;
}

/** Sets a solid-color outline on a shape's spPr. */
export const setSolidStroke = (spPr: XmlElement, options: StrokeOptions): void => {
  const ln = ensureLn(spPr);
  if (options.widthEmu !== undefined) {
    ln.attrs = ln.attrs.filter((a) => a.name.localName !== 'w');
    ln.attrs.push(attr(ATTR_W, String(Math.round(options.widthEmu))));
  }
  // Replace any existing fill choice inside <a:ln>.
  removeChildrenIn(ln, FILL_LOCALS);
  if (options.color !== undefined) {
    ln.children.unshift(elem(NAME_SOLID_FILL, { children: [buildColorElement(options.color)] }));
  }
};

/** Sets an explicit "no outline" on a shape's spPr. */
export const setNoStroke = (spPr: XmlElement): void => {
  const ln = ensureLn(spPr);
  removeChildrenIn(ln, FILL_LOCALS);
  ln.children.unshift(elem(NAME_NO_FILL));
};

/** Removes any `<a:ln>` from a shape's spPr (restores inheritance). */
export const clearStroke = (spPr: XmlElement): void => {
  spPr.children = spPr.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'ln'),
  );
};

/**
 * ECMA-376 §20.1.10.49 `ST_PresetLineDashVal` tokens. PowerPoint's
 * "Dash type" dropdown maps to these.
 */
export type LineDash =
  | 'solid'
  | 'dot'
  | 'dash'
  | 'lgDash'
  | 'dashDot'
  | 'lgDashDot'
  | 'lgDashDotDot'
  | 'sysDash'
  | 'sysDot'
  | 'sysDashDot'
  | 'sysDashDotDot';

const NAME_PRST_DASH = qname('a', 'prstDash', NS.dml);
const ATTR_VAL = qname('', 'val', '');

/**
 * Sets `<a:prstDash val="..."/>` inside the shape's `<a:ln>`. Creates
 * `<a:ln>` if absent. Replacing the dash choice on subsequent calls.
 */
export const setStrokeDash = (spPr: XmlElement, dash: LineDash): void => {
  const ln = ensureLn(spPr);
  ln.children = ln.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'prstDash'),
  );
  ln.children.push(elem(NAME_PRST_DASH, { attrs: [attr(ATTR_VAL, dash)] }));
};

/** ECMA-376 §20.1.10.39 `ST_LineEndType`. */
export type LineEndType = 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';

/** ECMA-376 `ST_LineEndWidth` / `ST_LineEndLength`. */
export type LineEndSize = 'sm' | 'med' | 'lg';

export interface ArrowOptions {
  readonly type: LineEndType;
  readonly width?: LineEndSize;
  readonly length?: LineEndSize;
}

const ATTR_TYPE = qname('', 'type', '');
const ATTR_LEN = qname('', 'len', '');

/**
 * Sets an arrowhead on one end of the shape's outline. `end` selects
 * which sentinel element: `'head'` → `<a:headEnd>`, `'tail'` →
 * `<a:tailEnd>`. Creates `<a:ln>` if absent and replaces any prior
 * arrowhead on the selected end.
 *
 * Pass `{ type: 'none' }` to clear an existing arrowhead (this still
 * emits the element with `type="none"`, matching PowerPoint's
 * behavior).
 */
export const setStrokeArrow = (
  spPr: XmlElement,
  end: 'head' | 'tail',
  options: ArrowOptions,
): void => {
  const ln = ensureLn(spPr);
  const localName = end === 'head' ? 'headEnd' : 'tailEnd';
  ln.children = ln.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
  );
  const attrs = [attr(ATTR_TYPE, options.type)];
  if (options.width !== undefined) attrs.push(attr(ATTR_W, options.width));
  if (options.length !== undefined) attrs.push(attr(ATTR_LEN, options.length));
  ln.children.push(elem(qname('a', localName, NS.dml), { attrs }));
};
