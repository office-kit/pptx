// Shape mutation: shadow + glow effects.

import { parseEffectList, resolveDrawingColor, resolveDrawingColorOpacity } from './shape-color.ts';
import { getShapePlaceholderIdx, getShapePlaceholderType } from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import { getEffectiveColorMap } from './color-map.ts';
import {
  type GlowOptions,
  type InnerShadowOptions,
  type ReadText3D,
  type ReflectionOptions,
  type ShadowOptions,
  type Text3D,
  applyShape3D,
  clearEffects as clearEffectsImpl,
  readText3D,
  removeEffect,
  setGlow,
  setInnerShadow,
  setReflection,
  setShadow,
  setSoftEdge,
} from '../../internal/drawingml/index.ts';
import { partName, resolveTarget } from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
} from '../../internal/xml/index.ts';

/** A shape's 3-D (`setShape3D`): the same DrawingML vocabulary as text 3-D. */
export type Shape3D = Text3D;
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, decode, requireSpPr } from './_helpers.ts';
import { getShapeStyleTheme, readShapeStyleEffectElement } from './shape-style-read.ts';
// ---------------------------------------------------------------------------
// Effects: shadow + glow.

const hasEffectSource = (shapeElement: XmlElement): boolean => {
  const spPr = firstChildElement(shapeElement, qname('p', 'spPr', NS.pml));
  if (!spPr) return false;
  if (
    firstChildElement(spPr, qname('a', 'effectLst', NS.dml)) !== null ||
    firstChildElement(spPr, qname('a', 'effectDag', NS.dml)) !== null
  )
    return true;
  const style = firstChildElement(shapeElement, qname('p', 'style', NS.pml));
  return style !== null && firstChildElement(style, qname('a', 'effectRef', NS.dml)) !== null;
};

/**
 * Read-back for `setShapeShadow` / `setShapeGlow`. Returns the kind
 * of effect currently on the shape's `<a:effectLst>`, or `null` when
 * none. Decodes the configured color + numeric parameters when
 * present.
 */
export type ShapeEffect =
  | {
      readonly kind: 'shadow';
      readonly color: string;
      readonly blurEmu: number;
      readonly offsetEmu: number;
      readonly angleDeg: number;
      readonly opacity?: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly radiusEmu: number;
    };

/**
 * Discriminated union covering every effect in
 * `CT_EffectStyleItem` (ECMA-376 §20.1.8.x) — outer shadow, inner
 * shadow, glow, reflection, soft-edge, blur. Returned in document
 * order so renderers can chain filters with the same composition
 * the reference desktop app applies.
 *
 * Lengths are EMU; angles are degrees clockwise from 3 o'clock;
 * opacity is a unit fraction (0..1) when the spec exposes one.
 */
export type ShapeEffectAny =
  | {
      readonly kind: 'outerShdw';
      readonly alignment?: ShadowOptions['alignment'];
      readonly rotateWithShape?: boolean;
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
      readonly scaleX?: number;
      readonly scaleY?: number;
      readonly skewX?: number;
      readonly skewY?: number;
    }
  | {
      readonly kind: 'innerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | {
      readonly kind: 'glow';
      readonly color: string;
      readonly opacity?: number;
      readonly radiusEmu: number;
    }
  | {
      readonly kind: 'reflection';
      // `opacity` is the far-end alpha (`endA`); `startOpacity` the
      // near-end alpha (`stA`). Both are unit fractions when authored.
      readonly opacity?: number;
      readonly startOpacity?: number;
      readonly startPosition?: number;
      readonly endPosition?: number;
      readonly fadeDirection?: number;
      readonly scaleX?: number;
      readonly alignment?: 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';
      readonly rotateWithShape?: boolean;
      // Vertical scale (`sy`) as a signed unit fraction — the reference desktop app
      // encodes the mirror as a negative `sy` (e.g. -1 = full-height
      // flip), so renderers must honor the sign, not just the magnitude.
      readonly scaleY?: number;
      readonly skewX?: number;
      readonly skewY?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
    }
  | { readonly kind: 'softEdge'; readonly radiusEmu: number }
  | { readonly kind: 'blur'; readonly radiusEmu: number };

export const getShapeEffect = (shape: SlideShapeData): ShapeEffect | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return null;

  const readColor = (host: XmlElement): { color: string; opacity?: number } => {
    const colorEl = host.children.find(
      (child): child is XmlElement =>
        child.kind === 'element' &&
        child.name.namespaceURI === NS.dml &&
        ['srgbClr', 'scrgbClr', 'hslClr', 'schemeClr', 'sysClr', 'prstClr'].includes(
          child.name.localName,
        ),
    );
    if (!colorEl) return { color: '' };
    const color = resolveDrawingColor(colorEl, null) ?? '';
    const opacity = resolveDrawingColorOpacity(colorEl);
    return { color, ...(opacity !== null ? { opacity } : {}) };
  };

  const outerShdw = firstChildElement(effectLst, qname('a', 'outerShdw', NS.dml));
  if (outerShdw) {
    const blur = Number.parseInt(getAttrValue(outerShdw, qname('', 'blurRad', '')) ?? '0', 10);
    const dist = Number.parseInt(getAttrValue(outerShdw, qname('', 'dist', '')) ?? '0', 10);
    const dirRaw = Number.parseInt(getAttrValue(outerShdw, qname('', 'dir', '')) ?? '0', 10);
    const c = readColor(outerShdw);
    return {
      kind: 'shadow',
      color: c.color,
      blurEmu: blur,
      offsetEmu: dist,
      angleDeg: dirRaw / 60000,
      ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
    };
  }
  const glow = firstChildElement(effectLst, qname('a', 'glow', NS.dml));
  if (glow) {
    const rad = Number.parseInt(getAttrValue(glow, qname('', 'rad', '')) ?? '0', 10);
    const c = readColor(glow);
    return { kind: 'glow', color: c.color, radiusEmu: rad };
  }
  return null;
};

/**
 * Returns every effect attached to the shape's `<a:effectLst>` in
 * document order — outer shadow, inner shadow, glow, reflection,
 * soft edge, blur. Empty array when no effects apply.
 *
 * Companion to `getShapeEffect`, which is the v1 "first effect only"
 * helper. `getShapeEffects` is what renderers want because the reference desktop app
 * composes multiple effects in a single filter (shadow + glow, etc.).
 */
export const getShapeEffects = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return [];
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (firstChildElement(spPr, qname('a', 'effectDag', NS.dml)) !== null) return [];
  // A shape style's effectRef is the theme-backed equivalent of an explicit
  // effectLst. An empty explicit effect list still wins, so only consult the
  // style reference when the shape has no local list at all.
  const resolved = effectLst ?? readShapeStyleEffectElement(pres, shape);
  if (!resolved) return [];
  return parseEffectList(
    resolved,
    getShapeStyleTheme(pres, shape).theme,
    getEffectiveColorMap(shape[SHAPE_SLIDE]),
  );
};

/**
 * Every shape on the slide whose `<a:effectLst>` carries an effect
 * of the given `kind` (`'outerShdw'`, `'innerShdw'`, `'glow'`,
 * `'reflection'`, `'softEdge'`, `'blur'`). Pure presence check — only
 * looks at the shape's own effect list, not the layout / master
 * cascade. Pair with `findShapesByEffect(pres, slide, 'softEdge')`
 * style call sites for visual-effect audits.
 */
export const findShapesByEffect = (
  pres: PresentationData,
  slide: SlideData,
  kind: ShapeEffectAny['kind'],
): ReadonlyArray<SlideShapeData> => {
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    if (getShapeEffects(pres, shape).some((e) => e.kind === kind)) out.push(shape);
  }
  return out;
};

/**
 * Same as `getShapeEffects` but walks the layout → master placeholder
 * cascade when the shape itself has no `<a:effectLst>`. Inherits
 * "all or nothing" — once any layer supplies an effect list, that
 * list is used; layers further down aren't merged in. This matches
 * the reference desktop app's behaviour (effect lists override rather than compose).
 */
export const getShapeEffectsEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const own = getShapeEffects(pres, shape);
  const ownEffectSource = hasEffectSource(shape[SHAPE_ELEMENT]);
  // An explicit empty effectLst, or effectRef idx="0", is a deliberate
  // clear and must stop placeholder inheritance just like a non-empty list.
  if (ownEffectSource) return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const theme = getShapeStyleTheme(pres, shape).theme;
  const layout = getSlideLayout(shape[SHAPE_SLIDE]);
  if (!layout) return own;

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

  const readEffectsOn = (el: XmlElement): readonly ShapeEffectAny[] | null => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return null;
    if (!hasEffectSource(el)) return null;
    // An authored effectDag or an unresolved effectRef is still an explicit
    // source. We cannot safely reinterpret it as inherited effects.
    if (firstChildElement(spPr, qname('a', 'effectDag', NS.dml)) !== null) return [];
    const eff =
      firstChildElement(spPr, qname('a', 'effectLst', NS.dml)) ??
      readShapeStyleEffectElement(pres, shape, el);
    return eff ? parseEffectList(eff, theme, getEffectiveColorMap(shape[SHAPE_SLIDE])) : [];
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const layoutEffects = readEffectsOn(layoutPh);
    if (layoutEffects !== null) return layoutEffects;
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
    const masterEffects = readEffectsOn(masterPh);
    if (masterEffects !== null) return masterEffects;
  }
  return own;
};

// Removing a shape's last effect drops its `<a:effectLst>`, except when the
// shape style references a theme effect: there an empty list is what keeps
// that effect from coming back.
const removeShapeEffect = (shape: SlideShapeData, localName: string): void => {
  const spPr = requireSpPr(shape);
  removeEffect(spPr, localName);
  const style = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'style', NS.pml));
  const effectRef = style && firstChildElement(style, qname('a', 'effectRef', NS.dml));
  const themed = effectRef !== null && getAttrValue(effectRef, qname('', 'idx', '')) !== '0';
  if (themed && firstChildElement(spPr, qname('a', 'effectLst', NS.dml)) === null) {
    // effectLst precedes scene3d, sp3d and extLst in CT_ShapeProperties.
    const at = spPr.children.findIndex(
      (c) =>
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        ['scene3d', 'sp3d', 'extLst'].includes(c.name.localName),
    );
    const list = elem(qname('a', 'effectLst', NS.dml), { children: [] });
    if (at === -1) spPr.children.push(list);
    else spPr.children.splice(at, 0, list);
  }
  commitAndRefresh(shape);
};

/**
 * Sets an outer drop shadow on the shape. Defaults: black, 4pt blur,
 * 3pt offset, 45° (down-right). Pass `opacity` (0–1) to soften the
 * shadow. Replaces an existing drop shadow and leaves the shape's other
 * effects in place; `null` removes only the drop shadow and
 * `clearShapeEffects` removes every effect.
 */
export const setShapeShadow = (shape: SlideShapeData, options: ShadowOptions | null = {}): void => {
  if (options === null) {
    removeShapeEffect(shape, 'outerShdw');
    return;
  }
  setShadow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets an inner shadow (`<a:innerShdw>`) on the shape, replacing only a
 * prior inner shadow. Defaults: black, 4pt blur, 3pt offset, 45°. `null`
 * removes it.
 */
export const setShapeInnerShadow = (
  shape: SlideShapeData,
  options: InnerShadowOptions | null,
): void => {
  if (options === null) {
    removeShapeEffect(shape, 'innerShdw');
    return;
  }
  setInnerShadow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a glow around the shape. The radius is in EMU (default 5pt =
 * 63500). Replaces an existing glow and composes with a shadow, in the
 * order `CT_EffectList` states. `null` removes only the glow.
 */
export const setShapeGlow = (shape: SlideShapeData, options: GlowOptions | null): void => {
  if (options === null) {
    removeShapeEffect(shape, 'glow');
    return;
  }
  setGlow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a reflection (`<a:reflection>`) of the shape, replacing only a prior
 * reflection. A mirrored reflection below the shape is `scaleY: -1` with
 * `alignment: 'bl'`, fading from `startOpacity` to `opacity` by
 * `endPosition`. `null` removes it.
 */
export const setShapeReflection = (
  shape: SlideShapeData,
  options: ReflectionOptions | null,
): void => {
  if (options === null) {
    removeShapeEffect(shape, 'reflection');
    return;
  }
  setReflection(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Softens the shape's edges (`<a:softEdge>`) by `radiusEmu`, replacing a
 * prior soft edge. `null` removes it.
 */
export const setShapeSoftEdge = (shape: SlideShapeData, radiusEmu: number | null): void => {
  if (radiusEmu === null) {
    removeShapeEffect(shape, 'softEdge');
    return;
  }
  setSoftEdge(requireSpPr(shape), radiusEmu);
  commitAndRefresh(shape);
};

/** Removes any effects (shadow / glow / future presets) from the shape. */
export const clearShapeEffects = (shape: SlideShapeData): void => {
  clearEffectsImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/**
 * Sets the shape's own 3-D — `<a:scene3d>` (camera, rotation, lighting) and
 * `<a:sp3d>` (bevels, depth, contour, material, distance from ground) in
 * `<p:spPr>`, what the reference desktop app's 3-D Format and 3-D Rotation write. Same
 * vocabulary as `setShapeText3D`, which puts it on the text body instead. A
 * field left out removes what it describes; settings this API does not model
 * (camera zoom, backdrop, ...) are kept while their element remains. `null`
 * removes both elements.
 */
export const setShape3D = (shape: SlideShapeData, value: Shape3D | null): void => {
  applyShape3D(requireSpPr(shape), value, 'setShape3D');
  commitAndRefresh(shape);
};

/**
 * Reads the 3-D on the shape's own `<p:spPr>` (see `setShape3D`), or `null`
 * when it has neither `<a:scene3d>` nor `<a:sp3d>`. A theme's effect style
 * is not consulted.
 */
export const getShape3D = (shape: SlideShapeData): ReadText3D | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  return spPr ? readText3D(spPr) : null;
};
