// Shape mutation: shadow + glow effects.

import { resolveDrawingColor } from './shape-color.ts';
import { getShapePlaceholderIdx, getShapePlaceholderType } from './shape-read-base.ts';
import { getSlideLayout } from './shape-slide-read.ts';
import {
  type GlowOptions,
  type ShadowOptions,
  clearEffects as clearEffectsImpl,
  setGlow,
  setShadow,
} from '../../internal/drawingml/index.ts';
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
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, decode, requireSpPr } from './_helpers.ts';
import { type PresentationTheme, getPresentationTheme } from './theme.ts';
// ---------------------------------------------------------------------------
// Effects: shadow + glow.

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
 * PowerPoint applies.
 *
 * Lengths are EMU; angles are degrees clockwise from 3 o'clock;
 * opacity is a unit fraction (0..1) when the spec exposes one.
 */
export type ShapeEffectAny =
  | {
      readonly kind: 'outerShdw';
      readonly color: string;
      readonly opacity?: number;
      readonly blurEmu: number;
      readonly distEmu: number;
      readonly angleDeg: number;
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
      readonly opacity?: number;
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
    const srgb = firstChildElement(host, qname('a', 'srgbClr', NS.dml));
    if (!srgb) return { color: '' };
    const val = getAttrValue(srgb, qname('', 'val', ''));
    const color = val !== null ? `#${val.toUpperCase()}` : '';
    const alpha = firstChildElement(srgb, qname('a', 'alpha', NS.dml));
    if (alpha) {
      const a = getAttrValue(alpha, qname('', 'val', ''));
      if (a !== null) {
        const n = Number.parseInt(a, 10);
        if (Number.isFinite(n)) return { color, opacity: n / 100000 };
      }
    }
    return { color };
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
 * helper. `getShapeEffects` is what renderers want because PowerPoint
 * composes multiple effects in a single filter (shadow + glow, etc.).
 */
// Parses an `<a:effectLst>` element into the typed effect union.
// Pulled out of `getShapeEffects` so the cascade-aware variant can
// reuse it.
const parseEffectLst = (
  effectLst: XmlElement,
  theme: PresentationTheme | null,
): ShapeEffectAny[] => {
  const readEffectColor = (host: XmlElement): { color: string; opacity?: number } => {
    let inner: XmlElement | null = null;
    for (const c of host.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      if (
        c.name.localName === 'srgbClr' ||
        c.name.localName === 'schemeClr' ||
        c.name.localName === 'sysClr' ||
        c.name.localName === 'prstClr'
      ) {
        inner = c;
        break;
      }
    }
    if (!inner) return { color: '' };
    let opacity: number | undefined;
    const alphaEl = firstChildElement(inner, qname('a', 'alpha', NS.dml));
    if (alphaEl) {
      const a = getAttrValue(alphaEl, qname('', 'val', ''));
      if (a !== null) {
        let n = Number.parseFloat(a);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
    }
    const hex = resolveDrawingColor(inner, theme);
    return { color: hex ?? '', ...(opacity !== undefined ? { opacity } : {}) };
  };

  const out: ShapeEffectAny[] = [];
  for (const child of effectLst.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.dml) continue;
    const local = child.name.localName;
    if (local === 'outerShdw' || local === 'innerShdw') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: local,
        color: c.color,
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'glow') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      const c = readEffectColor(child);
      out.push({
        kind: 'glow',
        color: c.color,
        radiusEmu: rad,
        ...(c.opacity !== undefined ? { opacity: c.opacity } : {}),
      });
    } else if (local === 'reflection') {
      const blur = Number.parseInt(getAttrValue(child, qname('', 'blurRad', '')) ?? '0', 10) || 0;
      const dist = Number.parseInt(getAttrValue(child, qname('', 'dist', '')) ?? '0', 10) || 0;
      const dir = Number.parseInt(getAttrValue(child, qname('', 'dir', '')) ?? '0', 10) || 0;
      const endA = getAttrValue(child, qname('', 'endA', ''));
      let opacity: number | undefined;
      if (endA !== null) {
        let n = Number.parseFloat(endA);
        if (Number.isFinite(n)) {
          if (Math.abs(n) > 1) n = n / 100000;
          opacity = n;
        }
      }
      out.push({
        kind: 'reflection',
        blurEmu: blur,
        distEmu: dist,
        angleDeg: dir / 60000,
        ...(opacity !== undefined ? { opacity } : {}),
      });
    } else if (local === 'softEdge') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'softEdge', radiusEmu: rad });
    } else if (local === 'blur') {
      const rad = Number.parseInt(getAttrValue(child, qname('', 'rad', '')) ?? '0', 10) || 0;
      out.push({ kind: 'blur', radiusEmu: rad });
    }
  }
  return out;
};

export const getShapeEffects = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return [];
  const effectLst = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
  if (!effectLst) return [];
  return parseEffectLst(effectLst, getPresentationTheme(pres));
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
 * PowerPoint's behaviour (effect lists override rather than compose).
 */
export const getShapeEffectsEffective = (
  pres: PresentationData,
  shape: SlideShapeData,
): readonly ShapeEffectAny[] => {
  const own = getShapeEffects(pres, shape);
  if (own.length > 0) return own;

  const phIdx = getShapePlaceholderIdx(shape);
  const phType = getShapePlaceholderType(shape);
  if (phIdx === null && phType === null) return own;

  const theme = getPresentationTheme(pres);
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

  const readEffectsOn = (el: XmlElement): readonly ShapeEffectAny[] => {
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    if (!spPr) return [];
    const eff = firstChildElement(spPr, qname('a', 'effectLst', NS.dml));
    if (!eff) return [];
    return parseEffectLst(eff, theme);
  };

  const layoutPh = findPh(layout[LAYOUT_PART].shapes);
  if (layoutPh) {
    const layoutEffects = readEffectsOn(layoutPh);
    if (layoutEffects.length > 0) return layoutEffects;
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
    if (masterEffects.length > 0) return masterEffects;
  }
  return own;
};

/**
 * Sets an outer drop shadow on the shape. Defaults: black, 4pt blur,
 * 3pt offset, 45° (down-right). Pass `opacity` (0–1) to soften the
 * shadow.
 */
export const setShapeShadow = (shape: SlideShapeData, options: ShadowOptions = {}): void => {
  setShadow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a glow around the shape. The radius is in EMU (default 5pt =
 * 63500). Mutually exclusive with `setShapeShadow` in v1 — calling
 * either replaces the prior `<a:effectLst>` entirely.
 */
export const setShapeGlow = (shape: SlideShapeData, options: GlowOptions): void => {
  setGlow(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/** Removes any effects (shadow / glow / future presets) from the shape. */
export const clearShapeEffects = (shape: SlideShapeData): void => {
  clearEffectsImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};
