// Detailed gradient-fill reader.

import { getShapePlaceholderIdx, getShapePlaceholderType } from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import { type GradientFillOptions } from '../../internal/drawingml/index.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { decode } from './_helpers.ts';
// ---------------------------------------------------------------------------
// Detailed gradient-fill reader. Companion to `getShapeFill`, which
// only reports the discriminated `kind`. Returns the full stop list +
// angle when the shape carries a `<a:gradFill>` of its own, or
// `null` for solid / pattern / image / none / inherited fills.
//
// Useful for renderers (preview generators, PDF exporters) that need
// to reproduce the gradient instead of substituting a placeholder.

export const NAME_A_GRAD_FILL = qname('a', 'gradFill', NS.dml);
export const NAME_A_GS_LST = qname('a', 'gsLst', NS.dml);
export const NAME_A_LIN = qname('a', 'lin', NS.dml);

export const readColorFromContainer = (parent: XmlElement): string | null => {
  for (const c of parent.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'srgbClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `#${val.toUpperCase()}`;
    }
    if (c.name.localName === 'schemeClr') {
      const val = getAttrValue(c, qname('', 'val', ''));
      if (val !== null) return `scheme:${val}`;
    }
  }
  return null;
};

// Parses one `<a:gradFill>` element into the stop list + direction.
// Shared by the shape-own reader and the placeholder-cascade reader.
const parseGradFill = (gradFill: XmlElement): GradientFillOptions | null => {
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs') {
      continue;
    }
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
  // ECMA-376 §20.1.8.33: gradFill has either <a:lin> (linear) or <a:path>
  // (radial / rectangular / shape-following) as a child to describe the
  // direction. We surface both so renderers can faithfully reproduce
  // non-linear gradients.
  let angleDeg = 0;
  const lin = firstChildElement(gradFill, NAME_A_LIN);
  if (lin) {
    const angRaw = getAttrValue(lin, qname('', 'ang', ''));
    if (angRaw !== null) {
      const ang = Number.parseInt(angRaw, 10);
      if (Number.isFinite(ang)) angleDeg = ang / 60_000;
    }
  }
  const pathEl = firstChildElement(gradFill, qname('a', 'path', NS.dml));
  if (pathEl) {
    const p = getAttrValue(pathEl, qname('', 'path', ''));
    const pathVal: 'circle' | 'rect' | 'shape' | null =
      p === 'circle' || p === 'rect' || p === 'shape' ? p : null;
    if (pathVal) {
      let focus: GradientFillOptions['focus'];
      const fillToRect = firstChildElement(pathEl, qname('a', 'fillToRect', NS.dml));
      if (fillToRect) {
        const pct = (name: string): number | undefined => {
          const v = getAttrValue(fillToRect, qname('', name, ''));
          if (v === null) return undefined;
          let n = Number.parseFloat(v);
          if (!Number.isFinite(n)) return undefined;
          if (Math.abs(n) > 1) n = n / 100000;
          return n;
        };
        const l = pct('l') ?? 0.5;
        const t = pct('t') ?? 0.5;
        const r = pct('r') ?? 0.5;
        const b = pct('b') ?? 0.5;
        focus = { left: l, top: t, right: r, bottom: b };
      }
      return { stops, angleDeg, path: pathVal, ...(focus ? { focus } : {}) };
    }
  }
  return { stops, angleDeg };
};

/**
 * Returns the full gradient definition (`stops` + `angleDeg`) when the
 * shape's `<p:spPr>` carries an `<a:gradFill>`. Returns `null` for any
 * other fill kind, including `inherit` — the function does not walk the
 * layout / master cascade. Use `getShapeGradientFillEffective` for that.
 */
export const getShapeGradientFill = (shape: SlideShapeData): GradientFillOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const gradFill = firstChildElement(spPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  return parseGradFill(gradFill);
};

/**
 * Same as `getShapeGradientFill` but walks the layout → master
 * placeholder cascade when the shape itself carries no `<a:gradFill>`.
 * Returns the first gradient found, or `null` when neither the shape
 * nor its inherited placeholder defines one.
 *
 * Resolves only gradients authored as a literal `<a:gradFill>` on the
 * shape or its placeholder ancestors. Gradients referenced through the
 * theme style matrix (`<p:style><a:fillRef>` → `<a:fillStyleLst>`) are
 * not modelled by the core yet, so those still report `inherit` and
 * fall through here.
 */
export const getShapeGradientFillEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): GradientFillOptions | null => {
  const own = getShapeGradientFill(shape);
  if (own) return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return null;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return null;

  const readGradFromSpPr = (el: XmlElement): GradientFillOptions | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    const gradFill = firstChildElement(spPr, NAME_A_GRAD_FILL);
    if (!gradFill) return null;
    return parseGradFill(gradFill);
  };

  const findPh = (
    shapes: ReadonlyArray<{
      placeholderIdx: number | null;
      placeholderType: string | null;
      element: XmlElement;
    }>,
  ): XmlElement | null => {
    let match = phIdx !== null ? shapes.find((s) => s.placeholderIdx === phIdx) : undefined;
    if (!match && phType !== null) match = shapes.find((s) => s.placeholderType === phType);
    return match?.element ?? null;
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const g = readGradFromSpPr(layoutPh);
    if (g) return g;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return null;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const g = readGradFromSpPr(masterPh);
    if (g) return g;
  }
  return null;
};
