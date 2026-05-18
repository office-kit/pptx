// Shape reads: fill and stroke.

import { resolveDrawingColor } from './shape-color.ts';
import { getShapePlaceholderIdx, getShapePlaceholderType } from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
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
import { getPresentationTheme } from './theme.ts';
export type ShapeFill =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Reads back the shape's stroke (`<a:ln>`). Returns:
 *
 *   - `{ kind: 'solid', color, widthEmu? }` for a solid-color outline.
 *   - `{ kind: 'none' }` when an `<a:noFill>` sits inside `<a:ln>`.
 *   - `{ kind: 'inherit' }` when no `<a:ln>` is present.
 */
export type ShapeStroke =
  | { readonly kind: 'solid'; readonly color: string; readonly widthEmu?: number }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Convenience over `getShapeStroke(shape)`: returns the solid-
 * stroke color (`#RRGGBB` / `scheme:<token>`) or `null` when the
 * stroke is inherited / removed.
 */
export const getShapeStrokeColor = (shape: SlideShapeData): string | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' ? stroke.color : null;
};

/**
 * Convenience over `getShapeStroke(shape)`: returns the stroke
 * width in EMU when the stroke is solid and an explicit width is
 * set, or `null` otherwise.
 */
export const getShapeStrokeWidth = (shape: SlideShapeData): number | null => {
  const stroke = getShapeStroke(shape);
  return stroke.kind === 'solid' && stroke.widthEmu !== undefined ? stroke.widthEmu : null;
};

/**
 * Returns the shape's stroke color resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the stroke isn't a solid color
 * (inherits / `noFill`) or when the color can't be resolved.
 *
 * Companion to `getShapeStrokeColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string — fine for round-tripping but
 * wrong for rendering, because PowerPoint paints the *transformed*
 * color, not the base one.
 */
export const getShapeStrokeColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const solid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Reads the stroke's line cap style — `'rnd'` (round), `'sq'` (square),
 * `'flat'`, or `null` when the attribute isn't set. Per ECMA-376
 * §20.1.2.3.10 (`ST_LineCap`).
 */
export const getShapeStrokeCap = (shape: SlideShapeData): 'rnd' | 'sq' | 'flat' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cap', ''));
  if (v === 'rnd' || v === 'sq' || v === 'flat') return v;
  return null;
};

/**
 * Reads the stroke's line join style — `'round'` / `'bevel'` / `'miter'`,
 * or `null` when no explicit join element is present. Maps from the
 * three child-element variants `<a:round/>`, `<a:bevel/>`, `<a:miter/>`.
 */
export const getShapeStrokeJoin = (shape: SlideShapeData): 'round' | 'bevel' | 'miter' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'round') return 'round';
    if (c.name.localName === 'bevel') return 'bevel';
    if (c.name.localName === 'miter') return 'miter';
  }
  return null;
};

/**
 * Reads the stroke's compound-line style (`<a:ln cmpd="…">`) — single,
 * double, triple, or thick/thin / thin/thick parallel lines. ECMA-376
 * §20.1.2.3.11 (`ST_CompoundLine`).
 */
export const getShapeStrokeCompound = (
  shape: SlideShapeData,
): 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const v = getAttrValue(ln, qname('', 'cmpd', ''));
  if (v === 'sng' || v === 'dbl' || v === 'thickThin' || v === 'thinThick' || v === 'tri') return v;
  return null;
};

/**
 * Same as `getShapeStroke` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. First non-inherit
 * stroke layer wins.
 */
export const getShapeStrokeEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): ShapeStroke => {
  const own = getShapeStroke(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readStrokeFromSpPr = (el: XmlElement): ShapeStroke | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
    if (!ln) return null;
    const wRaw = getAttrValue(ln, qname('', 'w', ''));
    const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;
    for (const c of ln.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (c.name.localName === 'noFill') return { kind: 'none' };
      if (c.name.localName === 'solidFill') {
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `#${val.toUpperCase()}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) {
              return {
                kind: 'solid',
                color: `scheme:${val}`,
                ...(widthEmu !== undefined ? { widthEmu } : {}),
              };
            }
          }
        }
      }
    }
    return null;
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
    const s = readStrokeFromSpPr(layoutPh);
    if (s) return s;
  }
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const s = readStrokeFromSpPr(masterPh);
    if (s) return s;
  }
  return own;
};

export const getShapeStroke = (shape: SlideShapeData): ShapeStroke => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return { kind: 'inherit' };
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return { kind: 'inherit' };

  const wRaw = getAttrValue(ln, qname('', 'w', ''));
  const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;

  for (const c of ln.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noFill') return { kind: 'none' };
    if (c.name.localName === 'solidFill') {
      for (const inner of c.children) {
        if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
        if (inner.name.localName === 'srgbClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `#${val.toUpperCase()}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
        if (inner.name.localName === 'schemeClr') {
          const val = getAttrValue(inner, qname('', 'val', ''));
          if (val !== null) {
            return {
              kind: 'solid',
              color: `scheme:${val}`,
              ...(widthEmu !== undefined ? { widthEmu } : {}),
            };
          }
        }
      }
      return {
        kind: 'solid',
        color: '',
        ...(widthEmu !== undefined ? { widthEmu } : {}),
      };
    }
  }
  return { kind: 'inherit' };
};

/**
 * Convenience over `getShapeFill(shape)`: returns the solid-fill
 * color string (`#RRGGBB` or `scheme:<token>`) when the shape has
 * one, or `null` otherwise. Use when the caller only cares about
 * the color and doesn't need to distinguish "inherit" / "no fill" /
 * "gradient" / "pattern" / "image" from each other.
 */
export const getShapeFillColor = (shape: SlideShapeData): string | null => {
  const fill = getShapeFill(shape);
  return fill.kind === 'solid' ? fill.color : null;
};

/**
 * Returns the shape's solid fill resolved to a concrete `#RRGGBB`:
 * scheme tokens are mapped through the deck's color scheme and
 * `<a:lumMod>` / `<a:tint>` / `<a:shade>` / etc. transform children
 * are applied. Returns `null` when the fill isn't solid (gradient,
 * pattern, image, none, inherit) or when the color can't be resolved.
 *
 * Companion to `getShapeFillColor`, which surfaces only the raw
 * `#RRGGBB` / `scheme:<token>` string. Renderers and exporters that
 * need the color PowerPoint actually paints should call this.
 */
export const getShapeFillColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const solid = firstChildElement(spPr, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
    return resolveDrawingColor(inner, getPresentationTheme(pres));
  }
  return null;
};

/**
 * Same as `getShapeFill` but walks the layout → master placeholder
 * cascade when the shape itself reports `'inherit'`. Returns the first
 * non-inherit fill found, or `{ kind: 'inherit' }` when neither layer
 * supplies one. Useful for renderers that want the actual fill the
 * placeholder will paint with.
 */
export const getShapeFillEffective = (pres: PresentationData, shape: SlideShapeData): ShapeFill => {
  const own = getShapeFill(shape);
  if (own.kind !== 'inherit') return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readFillFromSpPr = (el: XmlElement): ShapeFill | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    for (const c of spPr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      switch (c.name.localName) {
        case 'noFill':
          return { kind: 'none' };
        case 'solidFill': {
          for (const inner of c.children) {
            if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
            if (inner.name.localName === 'srgbClr') {
              const val = getAttrValue(inner, qname('', 'val', ''));
              if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
            }
            if (inner.name.localName === 'schemeClr') {
              const val = getAttrValue(inner, qname('', 'val', ''));
              if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
            }
          }
          return { kind: 'solid', color: '' };
        }
        case 'gradFill':
          return { kind: 'gradient' };
        case 'pattFill':
          return { kind: 'pattern' };
        case 'blipFill':
          return { kind: 'image' };
      }
    }
    return null;
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
    const f = readFillFromSpPr(layoutPh);
    if (f) return f;
  }

  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return own;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return own;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return own;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const { shapes: masterShapes } = readShapeTreeFromCsldRoot(masterRoot, 'sldMaster');
  const masterPh = findPh(masterShapes);
  if (masterPh) {
    const f = readFillFromSpPr(masterPh);
    if (f) return f;
  }
  return own;
};

export const getShapeFill = (shape: SlideShapeData): ShapeFill => {
  const spPrName = qname('p', 'spPr', NS.pml);
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], spPrName);
  if (!spPr) return { kind: 'inherit' };
  for (const c of spPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
      case 'noFill':
        return { kind: 'none' };
      case 'solidFill': {
        // Look for the immediate color choice; report sRGB verbatim,
        // scheme colors as "scheme:<token>".
        for (const inner of c.children) {
          if (inner.kind !== 'element' || inner.name.namespaceURI !== NS.dml) continue;
          if (inner.name.localName === 'srgbClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `#${val.toUpperCase()}` };
          }
          if (inner.name.localName === 'schemeClr') {
            const val = getAttrValue(inner, qname('', 'val', ''));
            if (val !== null) return { kind: 'solid', color: `scheme:${val}` };
          }
        }
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
    }
  }
  return { kind: 'inherit' };
};
