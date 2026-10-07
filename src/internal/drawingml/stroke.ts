// Outline (line) mutation for shapes.
//
// `<a:ln>` sits inside `<p:spPr>` after the fill choice. ECMA-376 §20.1.2
// surface: width (EMU), cap, dash, fill choice (solid/no/grad), join
// (round/bevel/miter), and head/tail arrow markers. We expose width, solid
// and gradient paint, noFill, preset dash, join, head/tail arrowheads and the
// Office 2021 sketched style — each inserted at its CT_LineProperties slot via
// LN_CHILD_RANK below.

import type { Color } from './color.ts';
import { LINE_DASHES } from '../enum-values.ts';
import { oneOf, lineWidthEmu } from '../bounds.ts';
import {
  NS,
  type XmlElement,
  type XmlNode,
  attr,
  elem,
  firstChildElement,
  insertChildByRank,
  qname,
} from '../xml/index.ts';
import { editSolidColor } from './color.ts';
import {
  buildGradientFill,
  type GradientFillOptions,
  type ReadGradientFill,
  validateGradientFillOptions,
} from './fill.ts';

const NAME_LN = qname('a', 'ln', NS.dml);
const NAME_SOLID_FILL = qname('a', 'solidFill', NS.dml);
const NAME_NO_FILL = qname('a', 'noFill', NS.dml);
const ATTR_W = qname('', 'w', '');

const FILL_LOCALS = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill']);

// CT_LineProperties (a:ln) is an xsd:sequence: fill choice, then the dash,
// then the join choice, then head/tail arrowheads, then extLst. Sub-element
// setters must drop their element at the mandated slot rather than push to the
// end, or combining (e.g.) a dash + arrowheads + join emits invalid order.
const LN_CHILD_RANK: Record<string, number> = {
  noFill: 0,
  solidFill: 0,
  gradFill: 0,
  pattFill: 0,
  prstDash: 1,
  custDash: 1,
  round: 2,
  bevel: 2,
  miter: 2,
  headEnd: 3,
  tailEnd: 4,
  extLst: 5,
};
const lnChildRank = (el: XmlElement): number =>
  el.name.namespaceURI === NS.dml ? (LN_CHILD_RANK[el.name.localName] ?? 99) : 99;
const insertLnChild = (ln: XmlElement, el: XmlElement): void =>
  insertChildByRank(ln, el, lnChildRank);

const removeChildrenIn = (host: XmlElement, names: ReadonlySet<string>): void => {
  host.children = host.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.dml && names.has(c.name.localName)),
  );
};

const ensureLn = (spPr: XmlElement): XmlElement =>
  firstChildElement(spPr, NAME_LN) ?? insertLn(spPr, elem(NAME_LN));

const insertLn = (spPr: XmlElement, ln: XmlElement): XmlElement => {
  // <a:ln> goes AFTER the fill choice and BEFORE effects / scene3d / sp3d /
  // extLst per the schema. We insert at the index of the first element that
  // belongs after `<a:ln>`; otherwise append.
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

/**
 * A non-solid line paint — the fill choice of `CT_LineProperties` other than
 * `<a:solidFill>` (which `color` spells) and `<a:noFill>`. Only gradients are
 * modelled; PowerPoint's Format pane offers no other kind for a line.
 */
export type LineFill = { readonly kind: 'gradient' } & GradientFillOptions;

/** A line paint read back from a deck; stop colors widen to `string`. */
export type ReadLineFill = { readonly kind: 'gradient' } & ReadGradientFill;

export interface StrokeOptions {
  /** Line color. Same accepted forms as `setFill`. */
  color?: Color;
  /** Line width in EMU. PowerPoint's default for a hairline is 9525 (0.75pt). */
  widthEmu?: number;
  /** Solid outline opacity, from 0 (transparent) to 1 (opaque). */
  opacity?: number;
  /**
   * Gradient paint (`<a:gradFill>` inside `<a:ln>`), PowerPoint's Gradient
   * line. It is the same fill choice as `color`, so passing both, or `fill`
   * with `opacity`, is rejected; stop opacities carry a gradient's alpha.
   */
  fill?: LineFill;
}

/** Updates the supplied outline properties, preserving omitted properties. */
export const setSolidStroke = (spPr: XmlElement, options: StrokeOptions): void => {
  const existing = firstChildElement(spPr, NAME_LN);
  const ln = existing ?? elem(NAME_LN);
  applyStroke(ln, options);
  if (!existing) insertLn(spPr, ln);
};

/** Rejects a stroke edit that would write two fill choices. */
export const validateStrokeFill = (
  options: { readonly color?: unknown; readonly opacity?: unknown; readonly fill?: LineFill },
  caller: string,
): void => {
  if (options.fill === undefined) return;
  if (options.color !== undefined || options.opacity !== undefined)
    throw new Error(`${caller}: fill is exclusive with color and opacity; pass one`);
  oneOf(options.fill.kind, ['gradient'], `${caller}: fill.kind`);
  validateGradientFillOptions(options.fill, `${caller}: fill`);
};

/**
 * The same edit on an `<a:ln>` the caller located — a run's outline lives in
 * `<a:rPr>`, whose child order is its own, so it cannot go through `ensureLn`.
 */
export const applyStroke = (
  ln: XmlElement,
  options: StrokeOptions,
  caller = 'setShapeStroke',
): void => {
  validateStrokeFill(options, caller);
  const previous = firstChildElement(ln, NAME_SOLID_FILL)?.children.find(
    (child) => child.kind === 'element',
  );
  const paint =
    options.fill !== undefined
      ? buildGradientFill(options.fill, `${caller}: fill`)
      : options.color !== undefined || options.opacity !== undefined
        ? elem(NAME_SOLID_FILL, { children: [editSolidColor(previous, options)] })
        : undefined;
  if (options.widthEmu !== undefined) {
    const width = lineWidthEmu(options.widthEmu, `${caller}: widthEmu`);
    ln.attrs = ln.attrs.filter((a) => a.name.localName !== 'w');
    ln.attrs.push(attr(ATTR_W, String(width)));
  }
  // Width-only edits preserve theme references, color transforms and noFill.
  if (paint) {
    removeChildrenIn(ln, FILL_LOCALS);
    insertLnChild(ln, paint);
  }
};

/** Sets an explicit "no outline" on a shape's spPr. */
export const setNoStroke = (spPr: XmlElement): void => {
  const ln = ensureLn(spPr);
  removeChildrenIn(ln, FILL_LOCALS);
  insertLnChild(ln, elem(NAME_NO_FILL));
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
  oneOf(dash, LINE_DASHES, 'setShapeStrokeDash: dash');
  const ln = ensureLn(spPr);
  ln.children = ln.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'prstDash'),
  );
  insertLnChild(ln, elem(NAME_PRST_DASH, { attrs: [attr(ATTR_VAL, dash)] }));
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
  oneOf(end, ['head', 'tail'], 'setShapeStrokeArrow: end');
  oneOf(
    options.type,
    ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'],
    'setShapeStrokeArrow: type',
  );
  if (options.width !== undefined)
    oneOf(options.width, ['sm', 'med', 'lg'], 'setShapeStrokeArrow: width');
  if (options.length !== undefined)
    oneOf(options.length, ['sm', 'med', 'lg'], 'setShapeStrokeArrow: length');
  const ln = ensureLn(spPr);
  const localName = { head: 'headEnd', tail: 'tailEnd' }[end];
  ln.children = ln.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === localName),
  );
  const attrs = [attr(ATTR_TYPE, options.type)];
  if (options.width !== undefined) attrs.push(attr(ATTR_W, options.width));
  if (options.length !== undefined) attrs.push(attr(ATTR_LEN, options.length));
  insertLnChild(ln, elem(qname('a', localName, NS.dml), { attrs }));
};

/** ECMA-376 §20.1.10.30 `ST_LineCap`. */
export type LineCap = 'rnd' | 'sq' | 'flat';

/** Sets `<a:ln cap="…"/>`. Pass `null` to clear the attribute. */
export const setStrokeCap = (spPr: XmlElement, cap: LineCap | null): void => {
  if (cap !== null) oneOf(cap, ['rnd', 'sq', 'flat'], 'setShapeStrokeCap: cap');
  const ln = ensureLn(spPr);
  ln.attrs = ln.attrs.filter((a) => !(a.name.namespaceURI === '' && a.name.localName === 'cap'));
  if (cap !== null) ln.attrs.push(attr(qname('', 'cap', ''), cap));
};

/** Three child-element variants of `<a:ln>`: round / bevel / miter. */
export type LineJoin = 'round' | 'bevel' | 'miter';
const JOIN_LOCALS = new Set(['round', 'bevel', 'miter']);

/** Replaces the join child of `<a:ln>`. Pass `null` to remove it. */
export const setStrokeJoin = (spPr: XmlElement, join: LineJoin | null): void => {
  if (join !== null) oneOf(join, ['round', 'bevel', 'miter'], 'setShapeStrokeJoin: join');
  const ln = ensureLn(spPr);
  removeChildrenIn(ln, JOIN_LOCALS);
  if (join !== null) {
    // Join sits after <a:prstDash> and before <a:headEnd> per CT_LineProperties.
    insertLnChild(ln, elem(qname('a', join, NS.dml)));
  }
};

/** ECMA-376 §20.1.10.31 `ST_CompoundLine`. */
export type LineCompound = 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri';

/** Sets `<a:ln cmpd="…"/>`. Pass `null` to clear the attribute. */
export const setStrokeCompound = (spPr: XmlElement, cmpd: LineCompound | null): void => {
  if (cmpd !== null)
    oneOf(cmpd, ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'], 'setShapeStrokeCompound: cmpd');
  const ln = ensureLn(spPr);
  ln.attrs = ln.attrs.filter((a) => !(a.name.namespaceURI === '' && a.name.localName === 'cmpd'));
  if (cmpd !== null) ln.attrs.push(attr(qname('', 'cmpd', ''), cmpd));
};

/**
 * PowerPoint's Sketched style presets ([MS-ODRAWXML] §2.38,
 * `EG_LineSketchType`), named after their `ask:lineSketch*` elements.
 */
export type LineSketch = 'curved' | 'freehand' | 'scribble';

const LINE_SKETCHES: readonly LineSketch[] = ['curved', 'freehand', 'scribble'];

// [MS-ODRAWXML] §2.38 sketchyshapes. Office stores the props in an
// `<a:ln><a:extLst><a:ext>` under this URI; the spec's tables do not list the
// URI, so it is taken from Office-written files.
const NS_ASK = 'http://schemas.microsoft.com/office/drawing/2018/sketchyshapes';
const SKETCH_EXT_URI = '{C807C97D-BFC1-408E-A445-0C87EB9F89A2}';
const NAME_EXT_LST = qname('a', 'extLst', NS.dml);
const NAME_EXT = qname('a', 'ext', NS.dml);
const ATTR_URI = qname('', 'uri', '');
const ATTR_SD = qname('', 'sd', '');
const NAME_SKETCH_PROPS = qname('ask', 'lineSketchStyleProps', NS_ASK);
const NAME_SKETCH_TYPE = qname('ask', 'type', NS_ASK);
const SKETCH_ELEMENT: Record<LineSketch, string> = {
  curved: 'lineSketchCurved',
  freehand: 'lineSketchFreehand',
  scribble: 'lineSketchScribble',
};
const GEOMETRY_LOCALS = new Set(['prstGeom', 'custGeom']);
const isGeometry = (c: XmlNode): c is XmlElement =>
  c.kind === 'element' && c.name.namespaceURI === NS.dml && GEOMETRY_LOCALS.has(c.name.localName);

const sketchExt = (ln: XmlElement): XmlElement | null => {
  const extLst = firstChildElement(ln, NAME_EXT_LST);
  return (
    extLst?.children.find(
      (c): c is XmlElement =>
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'ext' &&
        c.attrs.some((a) => a.name.localName === 'uri' && a.value === SKETCH_EXT_URI),
    ) ?? null
  );
};

const sketchProps = (ln: XmlElement): XmlElement | null => {
  const ext = sketchExt(ln);
  return ext && firstChildElement(ext, NAME_SKETCH_PROPS);
};

/** Reads the sketched style from an `<a:ln>`; `null` for none. */
export const readStrokeSketch = (ln: XmlElement): LineSketch | null => {
  const props = sketchProps(ln);
  const type = props && firstChildElement(props, NAME_SKETCH_TYPE);
  const choice = type?.children.find((c): c is XmlElement => c.kind === 'element');
  if (!choice || choice.name.namespaceURI !== NS_ASK) return null;
  return LINE_SKETCHES.find((sketch) => SKETCH_ELEMENT[sketch] === choice.name.localName) ?? null;
};

/**
 * True when the shape's geometry already is the sketched outline: Office
 * writes the hand-drawn path into `<p:spPr>` and keeps the original geometry
 * inside the sketch props, so a renderer must not roughen it a second time.
 */
export const hasSketchedGeometry = (ln: XmlElement): boolean =>
  Boolean(sketchProps(ln)?.children.some(isGeometry));

/**
 * Sets (or with `null` removes) the sketched style on `spPr`'s outline.
 *
 * Office replaces the shape geometry with the generated hand-drawn path and
 * keeps the original inside the props. This writer cannot generate that path,
 * so it moves any kept original back into `spPr` and leaves the roughening to
 * the renderer: afterwards `spPr` holds the shape's true outline. `seed` is
 * written for a new sketch only; an existing one keeps its seed.
 */
export const setStrokeSketch = (
  spPr: XmlElement,
  sketch: LineSketch | null,
  seed: number,
): void => {
  if (sketch !== null) oneOf(sketch, LINE_SKETCHES, 'setShapeStrokeSketch: sketch');
  const existingLn = firstChildElement(spPr, NAME_LN);
  if (sketch === null && !existingLn) return;
  const ln = existingLn ?? ensureLn(spPr);
  const extLst = firstChildElement(ln, NAME_EXT_LST);
  const ext = sketchExt(ln);
  const props = sketchProps(ln);
  const original = props?.children.find(isGeometry);
  if (props && original) {
    const index = spPr.children.findIndex(isGeometry);
    if (index >= 0) spPr.children.splice(index, 1, original);
    props.children = props.children.filter((c) => c !== original);
  }
  if (sketch === null) {
    if (ext && extLst) {
      extLst.children = extLst.children.filter((c) => c !== ext);
      if (!extLst.children.some((c) => c.kind === 'element'))
        ln.children = ln.children.filter((c) => c !== extLst);
    }
    return;
  }
  const type = elem(NAME_SKETCH_TYPE, {
    children: [elem(qname('ask', SKETCH_ELEMENT[sketch], NS_ASK))],
  });
  if (props) {
    // CT_LineSketchStyleProperties is a sequence: geometry, type, seed, extLst.
    const current = firstChildElement(props, NAME_SKETCH_TYPE);
    if (current) props.children.splice(props.children.indexOf(current), 1, type);
    else {
      const next = props.children.findIndex(
        (c) =>
          c.kind === 'element' && (c.name.localName === 'seed' || c.name.localName === 'extLst'),
      );
      props.children.splice(next < 0 ? props.children.length : next, 0, type);
    }
    return;
  }
  const created = elem(NAME_SKETCH_PROPS, {
    attrs: [attr(ATTR_SD, String(seed >>> 0))],
    prefixDecls: new Map([['ask', NS_ASK]]),
    children: [type],
  });
  if (ext) ext.children.push(created);
  else {
    const newExt = elem(NAME_EXT, { attrs: [attr(ATTR_URI, SKETCH_EXT_URI)], children: [created] });
    if (extLst) extLst.children.push(newExt);
    else insertLnChild(ln, elem(NAME_EXT_LST, { children: [newExt] }));
  }
};
