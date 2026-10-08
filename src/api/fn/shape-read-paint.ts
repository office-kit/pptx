// Shape reads: fill and stroke.

import { resolveDrawingColor, resolveDrawingColorOpacity } from './shape-color.ts';
import { parseGradFill, readColorFromContainer } from './shape-gradient-read.ts';
import type { ReadGradientFill } from '../../internal/drawingml/index.ts';
import { getShapePlaceholderIdx, getShapePlaceholderType } from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import {
  getShapeStyleTheme,
  readShapeStyleFill,
  readShapeStyleLineElement,
} from './shape-style-read.ts';
import { containingGroupFillElement } from './shape-group-paint.ts';
import { getEffectiveColorMap } from './color-map.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import {
  NS,
  cloneElement,
  elem,
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
  | { readonly kind: 'background' }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

/**
 * Reads back the shape's stroke (`<a:ln>`). Returns:
 *
 *   - `{ kind: 'solid', color, widthEmu? }` for a solid-color outline.
 *   - `{ kind: 'gradient', widthEmu? }` for a gradient line; read its stops
 *     with `getShapeStrokeGradient`.
 *   - `{ kind: 'none' }` when an `<a:noFill>` sits inside `<a:ln>`.
 *   - `{ kind: 'inherit' }` when no `<a:ln>` is present.
 */
export type ShapeStroke =
  | { readonly kind: 'solid'; readonly color: string; readonly widthEmu?: number }
  | { readonly kind: 'gradient'; readonly widthEmu?: number }
  | { readonly kind: 'none' }
  | { readonly kind: 'inherit' };

const readDirectFill = (element: XmlElement): ShapeFill | null => {
  const spPr =
    firstChildElement(element, qname('p', 'spPr', NS.pml)) ??
    firstChildElement(element, qname('p', 'grpSpPr', NS.pml));
  const choices = spPr ? spPr.children : element.name.namespaceURI === NS.dml ? [element] : [];
  for (const child of choices) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    switch (child.name.localName) {
      case 'noFill':
        return { kind: 'none' };
      case 'solidFill': {
        const color = readColorFromContainer(child);
        return { kind: 'solid', color: color ?? '' };
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

const hasDirectGroupFill = (shape: SlideShapeData): boolean => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  return Boolean(
    spPr?.children.some(
      (child) =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        child.name.localName === 'grpFill',
    ),
  );
};

// The DrawingML color element inside `<a:solidFill>` under `container`,
// or null when there is no solid fill there.
const solidFillColorElement = (container: XmlElement): XmlElement | null => {
  const solid = firstChildElement(container, qname('a', 'solidFill', NS.dml));
  if (!solid) return null;
  for (const inner of solid.children) {
    if (inner.kind === 'element' && inner.name.namespaceURI === NS.dml) return inner;
  }
  return null;
};

const fillColorElement = (shape: SlideShapeData): XmlElement | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  return spPr ? solidFillColorElement(spPr) : null;
};

const strokeColorElement = (shape: SlideShapeData, pres?: PresentationData): XmlElement | null => {
  const line = pres ? readShapeStrokeLineElement(pres, shape) : directStrokeLineElement(shape);
  return line ? solidFillColorElement(line) : null;
};

const directStrokeLineElement = (shape: SlideShapeData): XmlElement | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  return spPr ? firstChildElement(spPr, qname('a', 'ln', NS.dml)) : null;
};

const lineChildKey = (child: XmlElement): string => {
  // These choices occupy the same slot in CT_LineProperties. Replacing the
  // style choice as a group keeps a direct <a:noFill/> or gradient paint from
  // leaving the inherited solid paint behind.
  if (
    child.name.localName === 'noFill' ||
    child.name.localName === 'solidFill' ||
    child.name.localName === 'gradFill' ||
    child.name.localName === 'pattFill' ||
    child.name.localName === 'blipFill'
  ) {
    return 'fill';
  }
  if (
    child.name.localName === 'round' ||
    child.name.localName === 'bevel' ||
    child.name.localName === 'miter'
  ) {
    return 'join';
  }
  return child.name.localName;
};

/**
 * Returns the effective line properties after the shape's direct `<a:ln>`
 * overlays its style-matrix `lnRef`. DrawingML allows a direct line to carry
 * only one property (for example `w`), so treating that line as a complete
 * replacement drops the style's paint and dash settings.
 */
export const readShapeStrokeLineElement = (
  pres: PresentationData,
  shape: SlideShapeData,
): XmlElement | null => {
  const direct = directStrokeLineElement(shape);
  const style = readShapeStyleLineElement(pres, shape);
  if (!direct) return style;
  // idx=0 is represented by a standalone <a:noFill/>. Normalize it to line
  // properties before overlaying the direct line: a direct width alone must
  // retain the explicit no-line choice, while a direct paint can replace it.
  const styleLine =
    style?.name.localName === 'noFill'
      ? elem(qname('a', 'ln', NS.dml), { children: [cloneElement(style)] })
      : style;
  if (!styleLine) return direct;

  const merged = cloneElement(styleLine);
  const directAttrs = new Map(
    direct.attrs.map((attribute) => [
      `${attribute.name.namespaceURI}\u0000${attribute.name.localName}`,
      attribute,
    ]),
  );
  const mergedAttrKeys = new Set<string>();
  merged.attrs = merged.attrs.map((attribute) => {
    const key = `${attribute.name.namespaceURI}\u0000${attribute.name.localName}`;
    mergedAttrKeys.add(key);
    return directAttrs.get(key) ?? attribute;
  });
  for (const [key, attribute] of directAttrs) {
    if (!mergedAttrKeys.has(key)) merged.attrs.push(attribute);
  }
  const directChildren = new Map<string, XmlElement>();
  for (const child of direct.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    directChildren.set(lineChildKey(child), child);
  }
  const mergedChildKeys = new Set<string>();
  merged.children = merged.children.map((child) => {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) return child;
    const key = lineChildKey(child);
    mergedChildKeys.add(key);
    const replacement = directChildren.get(key);
    return replacement ? cloneElement(replacement) : child;
  });
  for (const [key, child] of directChildren) {
    if (!mergedChildKeys.has(key)) merged.children.push(cloneElement(child));
  }
  return merged;
};

const strokeLineElement = (shape: SlideShapeData, pres?: PresentationData): XmlElement | null =>
  pres ? readShapeStrokeLineElement(pres, shape) : directStrokeLineElement(shape);

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
 * wrong for rendering, because the reference desktop app paints the *transformed*
 * color, not the base one.
 */
export const getShapeStrokeColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const color = strokeColorElement(shape, pres);
  return color
    ? resolveDrawingColor(
        color,
        getShapeStyleTheme(pres, shape).theme,
        getEffectiveColorMap(shape[SHAPE_SLIDE]),
      )
    : null;
};

/**
 * Returns the opacity (`0`–`1`) of the shape's own solid outline, read from
 * the `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children of its color
 * element, or `null` when the outline isn't a solid color or carries no
 * alpha transform (the reference desktop app draws it fully opaque). Companion to
 * `getShapeStrokeColorResolved`, which never carries the alpha channel.
 */
export const getShapeStrokeOpacity = (
  shape: SlideShapeData,
  pres?: PresentationData,
): number | null => {
  const color = strokeColorElement(shape, pres);
  return color ? resolveDrawingColorOpacity(color) : null;
};

/**
 * Reads the stroke's line cap style — `'rnd'` (round), `'sq'` (square),
 * `'flat'`, or `null` when the attribute isn't set. Per ECMA-376
 * §20.1.2.3.10 (`ST_LineCap`).
 */
export const getShapeStrokeCap = (
  shape: SlideShapeData,
  pres?: PresentationData,
): 'rnd' | 'sq' | 'flat' | null => {
  const ln = strokeLineElement(shape, pres);
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
export const getShapeStrokeJoin = (
  shape: SlideShapeData,
  pres?: PresentationData,
): 'round' | 'bevel' | 'miter' | null => {
  const ln = strokeLineElement(shape, pres);
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
  pres?: PresentationData,
): 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null => {
  const ln = strokeLineElement(shape, pres);
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
  const line = effectiveStrokeLineElement(pres, shape);
  return (line && readStrokeElement(line)) ?? getShapeStroke(shape);
};

/**
 * The line element `getShapeStrokeEffective` reads: the shape's own line over
 * its style-matrix `lnRef`, else the first placeholder ancestor's line that
 * states a paint. `null` when no layer states one.
 */
const effectiveStrokeLineElement = (
  pres: PresentationData,
  shape: SlideShapeData,
): XmlElement | null => {
  const effectiveLine = readShapeStrokeLineElement(pres, shape);
  if (effectiveLine && readStrokeElement(effectiveLine)) return effectiveLine;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return null;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return null;

  // A placeholder ancestor's line counts once it states a usable paint; a
  // solid fill whose color cannot be read lets the cascade continue.
  const paintedLine = (el: XmlElement): XmlElement | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    const ln = spPr && firstChildElement(spPr, qname('a', 'ln', NS.dml));
    const stroke = ln && readStrokeElement(ln);
    return stroke && !(stroke.kind === 'solid' && stroke.color === '') ? ln : null;
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
  const fromLayout = layoutPh && paintedLine(layoutPh);
  if (fromLayout) return fromLayout;
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
  return masterPh && paintedLine(masterPh);
};

const readStrokeElement = (line: XmlElement): ShapeStroke | null => {
  if (line.name.localName === 'noFill') return { kind: 'none' };
  if (line.name.localName !== 'ln') return null;
  const wRaw = getAttrValue(line, qname('', 'w', ''));
  const widthEmu = wRaw !== null ? Number.parseInt(wRaw, 10) : undefined;
  const width = widthEmu !== undefined ? { widthEmu } : {};
  for (const c of line.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (c.name.localName === 'noFill') return { kind: 'none' };
    if (c.name.localName === 'solidFill') {
      const color = readColorFromContainer(c);
      return { kind: 'solid', color: color ?? '', ...width };
    }
    if (c.name.localName === 'gradFill') return { kind: 'gradient', ...width };
  }
  return null;
};

/**
 * Returns the full gradient (`stops`, direction, rotation) of a gradient line,
 * or `null` when the line is not one. Without `pres` only the shape's own
 * `<a:ln>` is read; with it, the line `getShapeStrokeEffective` resolves —
 * style matrix and placeholder cascade included — and every stop carries its
 * theme- and transform-resolved `resolvedColor`.
 */
export const getShapeStrokeGradient = (
  shape: SlideShapeData,
  pres?: PresentationData,
): ReadGradientFill | null => {
  const line = pres ? effectiveStrokeLineElement(pres, shape) : directStrokeLineElement(shape);
  const gradFill = line && firstChildElement(line, qname('a', 'gradFill', NS.dml));
  if (!gradFill) return null;
  return parseGradFill(
    gradFill,
    pres && {
      theme: getShapeStyleTheme(pres, shape).theme,
      colorMap: getEffectiveColorMap(shape[SHAPE_SLIDE]),
    },
  );
};

export const getShapeStroke = (shape: SlideShapeData): ShapeStroke => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return { kind: 'inherit' };
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return { kind: 'inherit' };
  return readStrokeElement(ln) ?? { kind: 'inherit' };
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
 * need the color the reference desktop app actually paints should call this.
 */
export const getShapeFillColorResolved = (
  pres: PresentationData,
  shape: SlideShapeData,
): string | null => {
  const color = fillColorElement(shape);
  if (color) return resolveDrawingColor(color, getPresentationTheme(pres));
  if (hasDirectGroupFill(shape)) {
    const groupFill = containingGroupFillElement(shape);
    const groupColor =
      groupFill?.name.localName === 'solidFill'
        ? groupFill.children.find(
            (child): child is XmlElement =>
              child.kind === 'element' && child.name.namespaceURI === NS.dml,
          )
        : null;
    if (!groupColor) return null;
    const { theme } = getShapeStyleTheme(pres, shape);
    return resolveDrawingColor(groupColor, theme, getEffectiveColorMap(shape[SHAPE_SLIDE]));
  }
  if (getShapeFill(shape).kind !== 'inherit') return null;
  const style = readShapeStyleFill(pres, shape);
  return style?.kind === 'solid' ? style.color || null : null;
};

/**
 * Returns the opacity (`0`–`1`) of the shape's solid fill, read from
 * the `<a:alpha>` / `<a:alphaMod>` / `<a:alphaOff>` children of its color
 * element, or `null` when the fill isn't solid or carries no alpha
 * transform (the reference desktop app paints it fully opaque). Companion to
 * `getShapeFillColorResolved`, which never carries the alpha channel —
 * OOXML encodes color and alpha independently. Pass `pres` to resolve the
 * shape's theme fill reference when no direct fill is present.
 */
export const getShapeFillOpacity = (
  shape: SlideShapeData,
  pres?: PresentationData,
): number | null => {
  const color = fillColorElement(shape);
  if (color) return resolveDrawingColorOpacity(color);
  if (hasDirectGroupFill(shape)) {
    const solid = containingGroupFillElement(shape);
    const color =
      solid?.name.localName === 'solidFill'
        ? solid.children.find(
            (child): child is XmlElement =>
              child.kind === 'element' && child.name.namespaceURI === NS.dml,
          )
        : null;
    return color ? resolveDrawingColorOpacity(color) : null;
  }
  if (!pres || getShapeFill(shape).kind !== 'inherit') return null;
  const style = readShapeStyleFill(pres, shape);
  return style?.kind === 'solid' ? resolveDrawingColorOpacity(style.colorElement) : null;
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

  // `grpFill` is an explicit DrawingML choice: the child paints with its
  // containing group's fill. It must not fall through to the child's style
  // matrix reference, which would incorrectly paint with the theme default.
  if (hasDirectGroupFill(shape))
    return readDirectFill(containingGroupFillElement(shape) ?? shape[SHAPE_ELEMENT]) ?? own;

  // A shape's own style reference supplies the default paint before the
  // placeholder layout/master cascade. Direct `spPr` paint above remains
  // authoritative, matching DrawingML's precedence rules.
  const style = readShapeStyleFill(pres, shape);
  if (style) {
    if (style.kind === 'solid') return { kind: 'solid', color: style.color };
    return style;
  }

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

  const readFillFromSpPr = (el: XmlElement): ShapeFill | null => {
    const background = getAttrValue(el, qname('', 'useBgFill', ''))?.trim();
    if (background === '1' || background === 'true') return { kind: 'background' };
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    for (const c of spPr.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      switch (c.name.localName) {
        case 'noFill':
          return { kind: 'none' };
        case 'solidFill': {
          const color = readColorFromContainer(c);
          if (color !== null) return { kind: 'solid', color };
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
  const background = getAttrValue(shape[SHAPE_ELEMENT], qname('', 'useBgFill', ''))?.trim();
  if (background === '1' || background === 'true') return { kind: 'background' };
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
        const color = readColorFromContainer(c);
        if (color !== null) return { kind: 'solid', color };
        return { kind: 'solid', color: '' };
      }
      case 'gradFill':
        return { kind: 'gradient' };
      case 'pattFill':
        return { kind: 'pattern' };
      case 'blipFill':
        return { kind: 'image' };
      case 'grpFill':
        return { kind: 'inherit' };
    }
  }
  return { kind: 'inherit' };
};
