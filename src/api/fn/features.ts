// Slide-level features: background, transition, speaker notes, and the
// presentation slide size.

import {
  type GradientFillOptions,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  setSolidFill,
  setTextBody,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  type ImageFormat,
  type PartName,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
  relsPartNameFor,
  resolveTarget,
} from '../../internal/opc/index.ts';
import {
  REL_TYPES,
  type TransitionOptions,
  buildEmptyNotesSlide,
  buildTransition,
  readPresentationPart,
} from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  parseXml,
  qname,
  serializeXml,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  LAYOUT_PART,
  LAYOUT_PART_NAME,
  type PresentationData,
  SLIDE_DOCUMENT,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideCommentData,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import {
  NAME_CSLD,
  NAME_SP_TREE,
  PRES_PART_NAME,
  commitSlideData,
  decode,
  encode,
  refreshSlideData,
  setOpcDefault,
} from './_helpers.ts';
import { getPresentationTheme } from './package.ts';
import { getSlides, isSlideHidden } from './slides.ts';
import {
  NAME_A_GRAD_FILL,
  NAME_A_GS_LST,
  NAME_A_LIN,
  type ShapeBounds,
  getShapeHyperlink,
  readColorFromContainer,
  resolveDrawingColor,
  setShapeHyperlink,
} from './shapes.ts';
import {
  type SlideChartData,
  findOverlappingShapePairs,
  getSlideCharts,
  getSlideComments,
  hasShapeImage,
  isChartShape,
  isTableShape,
} from './embedded.ts';

// ---------------------------------------------------------------------------
// Slide-level background + transition.

const removeTransition = (slide: SlideData): void => {
  slide[SLIDE_DOCUMENT].root.children = slide[SLIDE_DOCUMENT].root.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.pml &&
        c.name.localName === 'transition'
      ),
  );
};

const insertAfterClrMapOvr = (slide: SlideData, t: XmlElement): void => {
  const children = slide[SLIDE_DOCUMENT].root.children;
  let insertAt = children.length;
  for (let i = 0; i < children.length; i++) {
    const c = children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.pml) continue;
    if (c.name.localName === 'clrMapOvr') {
      insertAt = i + 1;
    } else if (c.name.localName === 'cSld' && insertAt === children.length) {
      insertAt = i + 1;
    }
  }
  children.splice(insertAt, 0, t);
};

/**
 * Reads back the slide's current transition (or `null` if no
 * `<p:transition>` is present). The returned shape mirrors what
 * `setSlideTransition` accepts.
 */
export const getSlideTransition = (slide: SlideData): TransitionOptions | null => {
  const transition = slide[SLIDE_DOCUMENT].root.children.find(
    (c): c is XmlElement =>
      c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'transition',
  );
  if (!transition) return null;
  const speed = getAttrValue(transition, qname('', 'spd', '')) as 'slow' | 'med' | 'fast' | null;
  const advClick = getAttrValue(transition, qname('', 'advClick', ''));
  const advTm = getAttrValue(transition, qname('', 'advTm', ''));
  // First child element identifies the effect (`p:fade`, `p:wipe`, ...).
  let effect: string | null = null;
  let direction: string | null = null;
  let orientation: 'horz' | 'vert' | null = null;
  let thruBlack: boolean | undefined;
  for (const child of transition.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    effect = child.name.localName;
    direction = getAttrValue(child, qname('', 'dir', ''));
    const o = getAttrValue(child, qname('', 'orient', ''));
    if (o === 'horz' || o === 'vert') orientation = o;
    const tb = getAttrValue(child, qname('', 'thruBlk', ''));
    if (tb !== null) thruBlack = tb === '1';
    break;
  }
  if (effect === null) return null;
  return {
    effect,
    ...(speed !== null ? { speed } : {}),
    ...(direction !== null ? { direction } : {}),
    ...(orientation !== null ? { orientation } : {}),
    ...(thruBlack !== undefined ? { thruBlack } : {}),
    ...(advClick !== null ? { advanceOnClick: advClick !== '0' } : {}),
    ...(advTm !== null ? { advanceAfterMs: Number.parseInt(advTm, 10) } : {}),
  };
};

/** Sets the slide's transition effect. */
export const setSlideTransition = (slide: SlideData, options: TransitionOptions): void => {
  removeTransition(slide);
  insertAfterClrMapOvr(slide, buildTransition(options));
  commitSlideData(slide);
  refreshSlideData(slide);
};

/** Removes any existing transition on the slide. */
export const clearSlideTransition = (slide: SlideData): void => {
  removeTransition(slide);
  commitSlideData(slide);
  refreshSlideData(slide);
};

const setSlideBackgroundXml = (slide: SlideData, configure: (bgPr: XmlElement) => void): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) throw new Error('slide has no <p:cSld>');
  const bgName = qname('p', 'bg', NS.pml);
  const bgPrName = qname('p', 'bgPr', NS.pml);
  let bg = firstChildElement(cSld, bgName);
  if (bg === null) {
    bg = { kind: 'element', name: bgName, attrs: [], prefixDecls: new Map(), children: [] };
    cSld.children.unshift(bg);
  }
  bg.children = [];
  const bgPr: XmlElement = {
    kind: 'element',
    name: bgPrName,
    attrs: [],
    prefixDecls: new Map(),
    children: [],
  };
  bg.children.push(bgPr);
  configure(bgPr);
  commitSlideData(slide);
  refreshSlideData(slide);
};

/**
 * Reads back the slide's current background. Returns a discriminated
 * union mirroring `getShapeFill`'s shape, plus `inherit` when no
 * `<p:bg>` element is present (the slide picks up its background from
 * the layout / master).
 */
export type SlideBackground =
  | { readonly kind: 'solid'; readonly color: string }
  | { readonly kind: 'gradient' }
  | { readonly kind: 'pattern' }
  | { readonly kind: 'image' }
  | { readonly kind: 'inherit' };

/**
 * Reads the slide's color-map override (`<p:clrMapOvr><p:overrideClrMapping/>`).
 * The mapping remaps the eight stable ECMA-376 color tokens (`bg1`,
 * `tx1`, `bg2`, `tx2`, `accent1`–`accent6`, `hlink`, `folHlink`) to
 * different theme positions. Returns `null` when the slide uses the
 * master's color map unchanged (the overwhelming common case).
 */
export const getSlideColorMapOverride = (slide: SlideData): Record<string, string> | null => {
  const root = slide[SLIDE_DOCUMENT].root;
  let ovr: XmlElement | null = null;
  for (const c of root.children) {
    if (c.kind !== 'element') continue;
    if (c.name.namespaceURI === NS.pml && c.name.localName === 'clrMapOvr') {
      ovr = c;
      break;
    }
  }
  if (!ovr) return null;
  const mapping = firstChildElement(ovr, qname('a', 'overrideClrMapping', NS.dml));
  if (!mapping) return null;
  // overrideClrMapping carries 12 attributes — bg1..folHlink — each
  // pointing to the index the token is remapped to in the theme.
  const out: Record<string, string> = {};
  for (const a of mapping.attrs) {
    if (a.name.namespaceURI !== '') continue;
    out[a.name.localName] = a.value;
  }
  return Object.keys(out).length > 0 ? out : null;
};

export const getSlideBackground = (slide: SlideData): SlideBackground => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  // <p:bg> can carry either a <p:bgPr> with explicit fill, or a
  // <p:bgRef idx="…"> that picks one of the theme's bgFillStyleLst /
  // fillStyleLst entries. The bgRef inner color element is the scheme
  // mapping target — projecting that to a scheme token is the most
  // useful shape for renderers.
  const bgRef = firstChildElement(bg, qname('p', 'bgRef', NS.pml));
  if (bgRef) {
    for (const inner of bgRef.children) {
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
    return { kind: 'inherit' };
  }
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
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
  return { kind: 'inherit' };
};

/**
 * A simplified, render-ready view of one of the layout's non-placeholder
 * shapes. Resolves the bounds, preset, fill, and stroke without going
 * through the slide-bound `SlideShapeData` symbols. Returned by
 * `getSlideLayoutBackgroundShapes` for renderers that want to paint the
 * layout's brand-template decoration (corner bars, divider lines, logos
 * as solid rects, etc.) behind the slide's own shapes.
 */
export interface SlideLayoutBackgroundShape {
  readonly kind: 'shape' | 'connector' | 'picture' | 'group' | 'graphicFrame';
  /** Bounds in EMU, or `null` when the shape inherits from its master. */
  readonly bounds: ShapeBounds | null;
  /** Preset geometry token for shapes (`'rect'`, `'roundRect'`, `'ellipse'`, …). */
  readonly preset: string | null;
  /** Fill color resolved to `#RRGGBB` (transforms + theme applied), or `null`. */
  readonly fillHex: string | null;
  /** Stroke color resolved to `#RRGGBB`, or `null`. */
  readonly strokeHex: string | null;
  /** Stroke width in EMU, or `null` when no explicit width is set. */
  readonly strokeWidthEmu: number | null;
  /** Rotation in degrees. */
  readonly rotation: number;
  /** Flip state. */
  readonly flip: { horizontal: boolean; vertical: boolean };
}

/**
 * Returns the non-placeholder shapes on a layout as a render-ready
 * view. Useful for previewing brand-template decoration (corner bars,
 * background rectangles, divider lines) that would otherwise be hidden
 * because they aren't reachable through `getSlideLayoutPlaceholders`.
 *
 * Placeholders are excluded — they're better rendered through the
 * slide's own placeholder bounds (which already cascade through the
 * layout). Picture and group shapes are omitted; their bytes / nested
 * children would need the layout's relationship table to resolve.
 */
export const getSlideLayoutBackgroundShapes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideLayoutBackgroundShape> => {
  const theme = getPresentationTheme(pres);
  const out: SlideLayoutBackgroundShape[] = [];
  for (const shape of layout[LAYOUT_PART].shapes) {
    if (shape.placeholderType !== null || shape.placeholderIdx !== null) continue;
    if (shape.kind !== 'shape' && shape.kind !== 'connector') continue;
    const el = shape.element;
    const pos = readPosition(el, shape.kind);
    const size = readSize(el, shape.kind);
    const bounds: ShapeBounds | null =
      pos !== null && size !== null
        ? { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu }
        : null;
    const spPr = firstChildElement(el, qname('p', 'spPr', NS.pml));
    let preset: string | null = null;
    let fillHex: string | null = null;
    let strokeHex: string | null = null;
    let strokeWidthEmu: number | null = null;
    if (spPr) {
      const prstGeom = firstChildElement(spPr, qname('a', 'prstGeom', NS.dml));
      if (prstGeom) preset = getAttrValue(prstGeom, qname('', 'prst', ''));
      const solid = firstChildElement(spPr, qname('a', 'solidFill', NS.dml));
      if (solid) {
        for (const c of solid.children) {
          if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
          fillHex = resolveDrawingColor(c, theme);
          break;
        }
      }
      const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
      if (ln) {
        const w = getAttrValue(ln, qname('', 'w', ''));
        if (w !== null) {
          const n = Number.parseInt(w, 10);
          if (Number.isFinite(n)) strokeWidthEmu = n;
        }
        const lnSolid = firstChildElement(ln, qname('a', 'solidFill', NS.dml));
        if (lnSolid) {
          for (const c of lnSolid.children) {
            if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
            strokeHex = resolveDrawingColor(c, theme);
            break;
          }
        }
      }
    }
    const rotation = readRotation(el, shape.kind);
    const flip = readFlip(el, shape.kind) ?? { horizontal: false, vertical: false };
    out.push({
      kind: shape.kind,
      bounds,
      preset,
      fillHex,
      strokeHex,
      strokeWidthEmu,
      rotation,
      flip,
    });
  }
  return out;
};

/**
 * Reads the slide layout's background. Same discriminated union as
 * `getSlideBackground` for slides — renderers fall back to this when
 * the slide's own background reports `'inherit'`. Walking one further
 * level to the master is left to callers (the same shape applies).
 */
/**
 * Reads the slide layout's pattern background when its `<p:bg>` is a
 * `<p:bgPr><a:pattFill>`. Same shape as `getSlideBackgroundPatternFill`
 * for slides.
 */
export const getSlideLayoutBackgroundPatternFill = (
  pres: PresentationData,
  layout: SlideLayoutData,
): { preset: string; foreground: string; background: string } | null => {
  const cSld = firstChildElement(layout[LAYOUT_PART].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const pattFill = firstChildElement(bgPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
};

/**
 * Reads the slide master's pattern background. Companion to
 * `getSlideLayoutBackgroundPatternFill`.
 */
export const getSlideMasterBackgroundPatternFill = (
  pres: PresentationData,
  layout: SlideLayoutData,
): { preset: string; foreground: string; background: string } | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return null;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const cSld = firstChildElement(masterRoot, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const pattFill = firstChildElement(bgPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
};

/**
 * Reads the slide layout's gradient background when its `<p:bg>` is a
 * `<p:bgPr><a:gradFill>`. Same shape as `getSlideBackgroundGradientFill`
 * for slides.
 */
export const getSlideLayoutBackgroundGradientFill = (
  layout: SlideLayoutData,
): GradientFillOptions | null => {
  const cSld = firstChildElement(layout[LAYOUT_PART].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const gradFill = firstChildElement(bgPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs')
      continue;
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
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
    if (pathVal) return { stops, angleDeg, path: pathVal };
  }
  return { stops, angleDeg };
};

/**
 * Reads the slide master's gradient background when present. Mirrors
 * `getSlideLayoutBackgroundGradientFill`; returns `null` for any other
 * background kind. Useful for closing the bg cascade — slides that
 * report `'gradient'` inherit can now get the master's gradient
 * projected via `gradientDef`.
 */
export const getSlideMasterBackgroundGradientFill = (
  pres: PresentationData,
  layout: SlideLayoutData,
): GradientFillOptions | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return null;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const cSld = firstChildElement(masterRoot, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const gradFill = firstChildElement(bgPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs')
      continue;
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
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
    if (pathVal) return { stops, angleDeg, path: pathVal };
  }
  return { stops, angleDeg };
};

/**
 * Reads the slide master's background. Same discriminated union as
 * `getSlideBackground` / `getSlideLayoutBackground`. Walks one rel up
 * from the layout to find the master, then reads `<p:bg>` (either
 * `<p:bgPr>` explicit fill or `<p:bgRef>` theme reference).
 *
 * Slides that report `'inherit'` on both their own background and the
 * layout's commonly inherit the brand fill from the master — this
 * closes the third level of the cascade.
 */
export const getSlideMasterBackground = (
  pres: PresentationData,
  layout: SlideLayoutData,
): SlideBackground => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return { kind: 'inherit' };
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return { kind: 'inherit' };
  const masterPart = pkg.getPart(resolveTarget(layoutPartName, masterRel.target));
  if (!masterPart) return { kind: 'inherit' };
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const cSld = firstChildElement(masterRoot, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  // bgRef on the master typically points at the theme's first
  // bgFillStyleLst entry; surface its inner color as a solid fill so
  // renderers paint the brand color.
  const bgRef = firstChildElement(bg, qname('p', 'bgRef', NS.pml));
  if (bgRef) {
    for (const inner of bgRef.children) {
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
    return { kind: 'inherit' };
  }
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
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
  return { kind: 'inherit' };
};

export const getSlideLayoutBackground = (layout: SlideLayoutData): SlideBackground => {
  const cSld = firstChildElement(layout[LAYOUT_PART].root, NAME_CSLD);
  if (!cSld) return { kind: 'inherit' };
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return { kind: 'inherit' };
  // <p:bgRef> = theme-reference fill (same shape as getSlideBackground).
  const bgRef = firstChildElement(bg, qname('p', 'bgRef', NS.pml));
  if (bgRef) {
    for (const inner of bgRef.children) {
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
    return { kind: 'inherit' };
  }
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return { kind: 'inherit' };
  for (const c of bgPr.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    switch (c.name.localName) {
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
  return { kind: 'inherit' };
};

/**
 * Returns the gradient stops + path when the slide carries a
 * `<p:bgPr><a:gradFill>` background. Returns `null` for any other
 * background kind. Shape identical to `getShapeGradientFill` so renderers
 * can use the same projection logic for slide backgrounds.
 */
export const getSlideBackgroundGradientFill = (slide: SlideData): GradientFillOptions | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const gradFill = firstChildElement(bgPr, NAME_A_GRAD_FILL);
  if (!gradFill) return null;
  // Reuse the same algorithm `getShapeGradientFill` does. The gradFill
  // element shape is identical between shape and slide backgrounds.
  const gsLst = firstChildElement(gradFill, NAME_A_GS_LST);
  if (!gsLst) return null;
  const stops: Array<{ offset: number; color: string }> = [];
  for (const c of gsLst.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml || c.name.localName !== 'gs')
      continue;
    const posRaw = getAttrValue(c, qname('', 'pos', ''));
    if (posRaw === null) continue;
    const pos = Number.parseInt(posRaw, 10);
    if (!Number.isFinite(pos)) continue;
    const color = readColorFromContainer(c);
    if (color === null) continue;
    stops.push({ offset: pos / 100_000, color });
  }
  if (stops.length === 0) return null;
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
    if (pathVal) return { stops, angleDeg, path: pathVal };
  }
  return { stops, angleDeg };
};

/**
 * Returns the pattern preset + theme-resolved colors when the slide
 * carries a `<p:bgPr><a:pattFill>` background. Returns `null` for any
 * other background kind. Shape mirrors `getShapePatternFill`.
 */
export const getSlideBackgroundPatternFill = (
  pres: PresentationData,
  slide: SlideData,
): { preset: string; foreground: string; background: string } | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const pattFill = firstChildElement(bgPr, qname('a', 'pattFill', NS.dml));
  if (!pattFill) return null;
  const preset = getAttrValue(pattFill, qname('', 'prst', '')) ?? 'pct50';
  const theme = getPresentationTheme(pres);
  const colorFrom = (parentName: string, fallback: string): string => {
    const parent = firstChildElement(pattFill, qname('a', parentName, NS.dml));
    if (!parent) return fallback;
    for (const c of parent.children) {
      if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
      const hex = resolveDrawingColor(c, theme);
      if (hex) return hex;
    }
    return fallback;
  };
  return {
    preset,
    foreground: colorFrom('fgClr', '#000000'),
    background: colorFrom('bgClr', '#FFFFFF'),
  };
};

/**
 * Returns the embedded image bytes when the slide carries a
 * `<p:bgPr><a:blipFill>` background, or `null` for any other background
 * kind (solid / gradient / pattern / inherit) or when the `r:embed`
 * relationship points at an external `r:link` target.
 *
 * Companion to `getSlideBackground`, which only reports the
 * discriminated `kind`. Renderers that want to actually paint the
 * background image (preview generators, snapshot tools) call this.
 */
export const getSlideBackgroundImageBytes = (slide: SlideData): Uint8Array | null => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const blipFill = firstChildElement(bgPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * Reads the slide layout's background image bytes when its `<p:bg>` is
 * a `<p:bgPr><a:blipFill>`. Same shape as `getSlideBackgroundImageBytes`
 * for slides. Returns `null` when the layout has no image background.
 */
export const getSlideLayoutBackgroundImageBytes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): Uint8Array | null => {
  const cSld = firstChildElement(layout[LAYOUT_PART].root, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const blipFill = firstChildElement(bgPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const rels = pkg.getRels(layoutPartName);
  if (!rels) return null;
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(layoutPartName, rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/**
 * Reads the slide master's background image bytes (via the layout's
 * master rel). Companion to `getSlideLayoutBackgroundImageBytes`.
 */
export const getSlideMasterBackgroundImageBytes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): Uint8Array | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutPartName = partName(layout[LAYOUT_PART_NAME]);
  const layoutRels = pkg.getRels(layoutPartName);
  if (!layoutRels) return null;
  const masterRel = layoutRels.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return null;
  const masterPartName = resolveTarget(layoutPartName, masterRel.target);
  const masterPart = pkg.getPart(masterPartName);
  if (!masterPart) return null;
  const masterRoot = parseXml(decode(masterPart.data)).root;
  const cSld = firstChildElement(masterRoot, NAME_CSLD);
  if (!cSld) return null;
  const bg = firstChildElement(cSld, qname('p', 'bg', NS.pml));
  if (!bg) return null;
  const bgPr = firstChildElement(bg, qname('p', 'bgPr', NS.pml));
  if (!bgPr) return null;
  const blipFill = firstChildElement(bgPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const masterRels = pkg.getRels(masterPartName);
  if (!masterRels) return null;
  const rel = masterRels.items.find((r) => r.id === rEmbed);
  if (!rel || rel.targetMode === 'External') return null;
  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(masterPartName, rel.target);
  const part = pkg.getPart(mediaName);
  return part?.data ?? null;
};

/** Sets a solid fill on the slide's background. */
export const setSlideBackground = (slide: SlideData, color: string): void => {
  setSlideBackgroundXml(slide, (bgPr) => setSolidFill(bgPr, color));
};

/**
 * Sets a picture as the slide's background. Embeds `bytes` as a new
 * media part, wires a slide → image rel, and replaces any prior
 * background with a `<p:bgPr><a:blipFill><a:blip r:embed="…"/>
 * <a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr>` payload.
 *
 * Format is detected from magic bytes; pass `options.format` to
 * override.
 */
export const setSlideBackgroundImage = (
  slide: SlideData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setSlideBackgroundImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const pkg = slide[INTERNAL_PACKAGE];

  // Allocate media part name.
  let nextN = 1;
  const mediaPattern = /^\/ppt\/media\/image(\d+)\./;
  for (const p of pkg.parts) {
    const m = p.name.match(mediaPattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const newMediaName = partName(`/ppt/media/image${nextN}.${extension}`);
  setOpcDefault(pkg, extension, contentType);
  pkg.addPart(newMediaName, contentType, bytes);

  // Slide → image rel.
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const newRId = nextRelId(rels.items.map((r) => r.id));
  rels.items.push({
    id: newRId,
    type: REL_TYPES.image,
    target: `../media/image${nextN}.${extension}`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], rels);

  setSlideBackgroundXml(slide, (bgPr) => {
    const blip = elem(qname('a', 'blip', NS.dml), {
      attrs: [attr(qname('r', 'embed', NS.officeDocRels), newRId)],
    });
    const stretch = elem(qname('a', 'stretch', NS.dml), {
      children: [elem(qname('a', 'fillRect', NS.dml))],
    });
    bgPr.children.push(elem(qname('a', 'blipFill', NS.dml), { children: [blip, stretch] }));
  });
};

/** Clears any explicit slide background, restoring layout inheritance. */
export const clearSlideBackground = (slide: SlideData): void => {
  const cSld = firstChildElement(slide[SLIDE_DOCUMENT].root, NAME_CSLD);
  if (!cSld) return;
  cSld.children = cSld.children.filter(
    (c) => !(c.kind === 'element' && c.name.namespaceURI === NS.pml && c.name.localName === 'bg'),
  );
  commitSlideData(slide);
  refreshSlideData(slide);
};

// ---------------------------------------------------------------------------
// Speaker notes.

const findNotesPartName = (slide: SlideData): PartName | null => {
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  const notesRel = rels.items.find((r) => r.type === REL_TYPES.notesSlide);
  if (!notesRel) return null;
  return notesRel.target.startsWith('/')
    ? partName(notesRel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], notesRel.target);
};

/**
 * Returns the slide's speaker notes (`null` if none). Pulls plain text
 * from the `body` placeholder; multi-line notes use `\n`.
 */
export const getSlideNotes = (slide: SlideData): string | null => {
  const notesPartName = findNotesPartName(slide);
  if (notesPartName === null) return null;
  const part = slide[INTERNAL_PACKAGE].getPart(notesPartName);
  if (part === null) return null;
  const root = parseXml(decode(part.data)).root;
  const cSld = firstChildElement(root, NAME_CSLD);
  if (!cSld) return null;
  const spTree = firstChildElement(cSld, NAME_SP_TREE);
  if (!spTree) return null;
  for (const child of spTree.children) {
    if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
    if (child.name.localName !== 'sp') continue;
    const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
    if (!nvSpPr) continue;
    const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
    if (!nvPr) continue;
    const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
    if (!ph) continue;
    const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
    if (!txBody) continue;
    const lines: string[] = [];
    for (const p of txBody.children) {
      if (p.kind !== 'element' || p.name.namespaceURI !== NS.dml || p.name.localName !== 'p') {
        continue;
      }
      let line = '';
      for (const r of p.children) {
        if (r.kind !== 'element' || r.name.namespaceURI !== NS.dml || r.name.localName !== 'r') {
          continue;
        }
        for (const tElement of r.children) {
          if (
            tElement.kind === 'element' &&
            tElement.name.namespaceURI === NS.dml &&
            tElement.name.localName === 't'
          ) {
            for (const tc of tElement.children) {
              if (tc.kind === 'text') line += tc.data;
            }
          }
        }
      }
      lines.push(line);
    }
    return lines.join('\n');
  }
  return null;
};

/**
 * Sets the slide's speaker notes. Creates the `notesSlide` part and
 * wires up the rels (slide ↔ notesSlide ↔ notesMaster) on first call;
 * subsequent calls just replace the body placeholder text.
 */
/**
 * Returns every slide whose `<p:sldId show="0">` flag is set —
 * complement of `getVisibleSlides`. Useful for audit UIs and
 * batch-unhide operations.
 */
export const getHiddenSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (isSlideHidden(slide)) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide in document order whose `<p:sldId show="0">`
 * flag is *not* set. Convenience over `getSlides(pres).filter(s =>
 * !isSlideHidden(s))` — useful when an export pipeline needs to
 * skip hidden slides without touching the `show` attribute itself.
 */
export const getVisibleSlides = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (!isSlideHidden(slide)) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one chart graphic frame.
 * Built on `isChartShape`.
 */
export const getSlidesWithCharts = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => isChartShape(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide where at least two shapes have overlapping
 * bounding boxes. Built on `findOverlappingShapePairs`. Useful for
 * deck-wide layout audits — surfacing slides that may have stacked
 * or accidentally-colliding content for human review.
 */
export const getSlidesWithOverlap = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (findOverlappingShapePairs(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one table graphic frame.
 * Built on `isTableShape`.
 */
export const getSlidesWithTables = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => isTableShape(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide carrying at least one image-bearing shape
 * (a `<p:pic>` picture or a regular shape with `<a:blipFill>`).
 * Built on `hasShapeImage`.
 */
export const getSlidesWithImages = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => hasShapeImage(s))) out.push(slide);
  }
  return out;
};

/**
 * Returns every slide that has at least one comment attached.
 * Convenience over `getSlides(pres).filter(s =>
 * getSlideComments(s).length > 0)`.
 */
export const getSlidesWithComments = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (getSlideComments(slide).length > 0) out.push(slide);
  }
  return out;
};

/**
 * One entry per slide with non-empty notes, carrying its 0-based
 * slide index and the notes text. Useful for "export speaker
 * notes to a separate document" workflows that need both the
 * notes and the slide they belong to.
 */
export interface PresentationNotesEntry {
  readonly slideIndex: number;
  readonly notes: string;
}

/**
 * One entry per chart in the deck, carrying both the chart and the
 * 0-based slide it was attached to.
 */
export interface PresentationChartEntry {
  readonly slideIndex: number;
  readonly chart: SlideChartData;
}

/**
 * One entry per image-bearing shape in the deck, carrying the
 * shape (picture or image-filled) and the 0-based slide it lives
 * on. Sibling of `getAllCharts` / `getAllTables`.
 */
export interface PresentationImageEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
}

/**
 * Returns every image-bearing shape across the deck (pictures and
 * shapes with `<a:blipFill>`), paired with its 0-based slide
 * index. Built on `hasShapeImage`.
 */
export const getAllImages = (pres: PresentationData): ReadonlyArray<PresentationImageEntry> => {
  const out: PresentationImageEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      if (hasShapeImage(shape)) out.push({ slideIndex: i, shape });
    }
  }
  return out;
};

/**
 * One entry per table in the deck, carrying the table shape and
 * the 0-based slide it sits on. Sibling of `getAllCharts`.
 */
export interface PresentationTableEntry {
  readonly slideIndex: number;
  readonly table: SlideShapeData;
}

/**
 * Returns every table across every slide in the deck, paired with
 * the 0-based index of its slide. Built on `isTableShape`.
 */
export const getAllTables = (pres: PresentationData): ReadonlyArray<PresentationTableEntry> => {
  const out: PresentationTableEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      if (isTableShape(shape)) out.push({ slideIndex: i, table: shape });
    }
  }
  return out;
};

/**
 * Returns every chart across every slide in the deck, paired with
 * the 0-based index of its slide. Useful for chart-inventory UIs
 * and bulk chart-update pipelines.
 */
export const getAllCharts = (pres: PresentationData): ReadonlyArray<PresentationChartEntry> => {
  const out: PresentationChartEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const c of getSlideCharts(slides[i]!)) {
      out.push({ slideIndex: i, chart: c });
    }
  }
  return out;
};

/**
 * One entry per external hyperlink found in a shape's text body,
 * carrying the URL, the linked shape, and the 0-based slide index.
 * Each hyperlinked shape is reported once (the URL of its first
 * `<a:hlinkClick>` run).
 */
export interface PresentationHyperlinkEntry {
  readonly slideIndex: number;
  readonly shape: SlideShapeData;
  readonly url: string;
}

/**
 * Returns every external hyperlink in the deck — one entry per
 * shape whose text body carries an `<a:hlinkClick>`. Useful for
 * "link audit" passes before publishing, and for building a
 * deck-wide table of contents of outbound URLs.
 */
export const getAllHyperlinks = (
  pres: PresentationData,
): ReadonlyArray<PresentationHyperlinkEntry> => {
  const out: PresentationHyperlinkEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const shape of slides[i]![SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url !== null) out.push({ slideIndex: i, shape, url });
    }
  }
  return out;
};

/**
 * Returns every distinct external URL referenced by any shape in
 * the deck, in first-seen order. Sibling of `getAllHyperlinks`
 * (which keeps duplicates and slide indices). Useful for "are
 * these URLs all live?" audits where checking each URL once is
 * enough.
 */
export const getDistinctHyperlinkUrls = (pres: PresentationData): ReadonlyArray<string> => {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url !== null && !seen.has(url)) {
        seen.add(url);
        out.push(url);
      }
    }
  }
  return out;
};

/**
 * Returns every slide carrying at least one shape with an external
 * hyperlink. Built on `findHyperlinkedShapes`. Useful for navigation
 * UIs that want to surface only the slides containing outbound URLs.
 */
export const getSlidesWithHyperlinks = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    if (slide[SLIDE_SHAPES].some((s) => getShapeHyperlink(s) !== null)) {
      out.push(slide);
    }
  }
  return out;
};

/**
 * Returns every slide containing at least one shape whose external
 * hyperlink matches `needle` (substring or `RegExp`). Sibling of
 * `findSlidesByText` for outbound-URL audits — e.g. "every slide
 * that links to old.docs.example.com".
 */
export const findSlidesByHyperlink = (
  pres: PresentationData,
  needle: string | RegExp,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      const url = getShapeHyperlink(shape);
      if (url === null) continue;
      const hit = typeof needle === 'string' ? url.includes(needle) : needle.test(url);
      if (hit) {
        out.push(slide);
        break;
      }
    }
  }
  return out;
};

/**
 * Bulk URL migration. Re-points every shape across the deck whose
 * first hyperlink exactly equals `from` to instead point at `to`.
 * Returns the number of shapes updated. Built on
 * `setShapeHyperlink`, so each update goes through the standard
 * rels-allocation path and stays schema-valid.
 *
 * Matching is exact (case-sensitive). To migrate by pattern, use
 * `findSlidesByHyperlink` to locate slides and rewrite each shape
 * yourself.
 */
export const replaceHyperlink = (pres: PresentationData, from: string, to: string): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      if (getShapeHyperlink(shape) === from) {
        setShapeHyperlink(shape, to);
        n++;
      }
    }
  }
  return n;
};

/**
 * Removes every external hyperlink across the deck — useful for
 * sanitizing a template before sharing, or for stripping outbound
 * URLs from an exported preview. Returns the number of shapes
 * cleared. Each call goes through `setShapeHyperlink(_, null)`.
 */
export const clearAllHyperlinks = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    for (const shape of slide[SLIDE_SHAPES]) {
      if (getShapeHyperlink(shape) !== null) {
        setShapeHyperlink(shape, null);
        n++;
      }
    }
  }
  return n;
};

/**
 * Slide-scoped sibling of `clearAllHyperlinks`. Removes every
 * external hyperlink on this slide and returns the number of
 * shapes cleared.
 */
export const clearSlideHyperlinks = (slide: SlideData): number => {
  let n = 0;
  for (const shape of slide[SLIDE_SHAPES]) {
    if (getShapeHyperlink(shape) !== null) {
      setShapeHyperlink(shape, null);
      n++;
    }
  }
  return n;
};

/**
 * One entry per comment in the deck, carrying both the comment and
 * the 0-based slide it was attached to.
 */
export interface PresentationCommentEntry {
  readonly slideIndex: number;
  readonly comment: SlideCommentData;
}

/**
 * Returns every comment across every slide in the deck, each paired
 * with the 0-based index of its slide. Useful for review-summary
 * UIs that show all annotations in one chronological list.
 */
export const getAllComments = (pres: PresentationData): ReadonlyArray<PresentationCommentEntry> => {
  const out: PresentationCommentEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    for (const c of getSlideComments(slides[i]!)) {
      out.push({ slideIndex: i, comment: c });
    }
  }
  return out;
};

/**
 * Returns every slide's speaker notes alongside its 0-based index.
 * Skips slides whose notes are empty / unset.
 */
export const getAllNotes = (pres: PresentationData): ReadonlyArray<PresentationNotesEntry> => {
  const out: PresentationNotesEntry[] = [];
  const slides = getSlides(pres);
  for (let i = 0; i < slides.length; i++) {
    const notes = getSlideNotes(slides[i]!);
    if (notes !== null && notes.length > 0) out.push({ slideIndex: i, notes });
  }
  return out;
};

/**
 * Returns every slide in the presentation that carries non-empty
 * speaker notes. Convenience over `getSlides(pres).filter(s =>
 * getSlideNotes(s) !== null && getSlideNotes(s) !== '')`.
 */
export const getSlidesWithNotes = (pres: PresentationData): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const notes = getSlideNotes(slide);
    if (notes !== null && notes.length > 0) out.push(slide);
  }
  return out;
};

/**
 * Predicate sibling of `getSlideNotes`. Returns `true` when the
 * slide carries a non-empty `notesSlide` body — i.e. whatever
 * `getSlideNotes(slide)` would surface is a non-empty string.
 *
 * Cheap to call in hot loops where the caller only needs to know
 * "are there any notes here?" without materializing the text.
 */
export const hasSlideNotes = (slide: SlideData): boolean => {
  const notes = getSlideNotes(slide);
  return notes !== null && notes.length > 0;
};

/**
 * Code-point length of the slide's speaker notes, or `0` when the
 * slide has no notes. Counts via `Array.from`, so surrogate-pair
 * characters (emoji, supplementary CJK) count as 1 — matches
 * `getSlideTextLength`.
 */
export const getSlideNotesLength = (slide: SlideData): number => {
  const notes = getSlideNotes(slide);
  return notes === null ? 0 : Array.from(notes).length;
};

/**
 * Concatenated speaker notes from every slide, joined with the
 * given `separator` (defaults to a form-feed, `\f`). Slides with
 * no notes contribute the empty string. Sibling of
 * `getPresentationText` — useful for search-indexing notes
 * across a whole deck.
 */
export const getPresentationNotesText = (
  pres: PresentationData,
  separator: string = '\f',
): string => {
  const parts: string[] = [];
  for (const slide of getSlides(pres)) parts.push(getSlideNotes(slide) ?? '');
  return parts.join(separator);
};

/**
 * Total code-point length of speaker notes across every slide.
 * Sibling of `getPresentationTextLength`; counts surrogate-pair
 * characters as 1 each. Cheaper than `getPresentationNotesText`
 * when the caller only needs the size.
 */
export const getPresentationNotesLength = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) n += getSlideNotesLength(slide);
  return n;
};

/**
 * Appends `text` to the slide's existing notes on its own line.
 * Equivalent to `setSlideNotes(slide, (getSlideNotes(slide) ?? '') + '\n' + text)`,
 * minus the leading newline when there were no notes yet.
 */
export const appendSlideNotes = (slide: SlideData, text: string): void => {
  const existing = getSlideNotes(slide);
  const value = existing === null || existing.length === 0 ? text : `${existing}\n${text}`;
  setSlideNotes(slide, value);
};

export const setSlideNotes = (slide: SlideData, value: string): void => {
  const pkg = slide[INTERNAL_PACKAGE];
  const notesPartName = findNotesPartName(slide);
  if (notesPartName !== null) {
    const part = pkg.getPart(notesPartName);
    if (part === null) throw new Error(`notes rel points at missing part ${notesPartName}`);
    const doc = parseXml(decode(part.data));
    const cSld = firstChildElement(doc.root, NAME_CSLD);
    if (!cSld) throw new Error('notesSlide has no <p:cSld>');
    const spTree = firstChildElement(cSld, NAME_SP_TREE);
    if (!spTree) throw new Error('notesSlide has no <p:spTree>');
    for (const child of spTree.children) {
      if (child.kind !== 'element' || child.name.namespaceURI !== NS.pml) continue;
      if (child.name.localName !== 'sp') continue;
      const nvSpPr = firstChildElement(child, qname('p', 'nvSpPr', NS.pml));
      if (!nvSpPr) continue;
      const nvPr = firstChildElement(nvSpPr, qname('p', 'nvPr', NS.pml));
      if (!nvPr) continue;
      const ph = firstChildElement(nvPr, qname('p', 'ph', NS.pml));
      if (!ph) continue;
      const txBody = firstChildElement(child, qname('p', 'txBody', NS.pml));
      if (!txBody) continue;
      setTextBody(txBody, value);
      part.data = encode(serializeXml(doc));
      return;
    }
    throw new Error('notesSlide has no body placeholder to fill');
  }

  // Create a new notesSlide part.
  const notesMasterPart = pkg.parts.find((p) => p.contentType.endsWith('notesMaster+xml'));
  let nextN = 1;
  const pattern = /^\/ppt\/notesSlides\/notesSlide(\d+)\.xml$/;
  for (const p of pkg.parts) {
    const m = p.name.match(pattern);
    if (m?.[1] !== undefined) {
      const n = Number.parseInt(m[1], 10);
      if (Number.isFinite(n) && n >= nextN) nextN = n + 1;
    }
  }
  const notesName = partName(`/ppt/notesSlides/notesSlide${nextN}.xml`);
  const doc = buildEmptyNotesSlide(value);
  pkg.addPart(
    notesName,
    'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    encode(serializeXml(doc)),
  );

  const notesRels = emptyRels();
  const slideBase = slide[SLIDE_PART_NAME].split('/').pop() ?? 'slide.xml';
  notesRels.items.push({
    id: 'rId1',
    type: REL_TYPES.slide,
    target: `../slides/${slideBase}`,
    targetMode: 'Internal',
  });
  if (notesMasterPart) {
    const notesMasterBase = notesMasterPart.name.split('/').pop() ?? 'notesMaster1.xml';
    notesRels.items.push({
      id: 'rId2',
      type: REL_TYPES.notesMaster,
      target: `../notesMasters/${notesMasterBase}`,
      targetMode: 'Internal',
    });
  }
  pkg.setRels(notesName, notesRels);

  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]) ?? emptyRels();
  const existingIds = slideRels.items.map((r) => r.id);
  let n = 1;
  while (existingIds.includes(`rId${n}`)) n++;
  slideRels.items.push({
    id: `rId${n}`,
    type: REL_TYPES.notesSlide,
    target: `../notesSlides/notesSlide${nextN}.xml`,
    targetMode: 'Internal',
  });
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
};

/**
 * Removes the slide's speaker-notes part entirely. Drops the
 * `notesSlide` part + its `.rels`, and unwires the slide → notesSlide
 * relationship. No-op when the slide has no notes.
 *
 * The shared `notesMaster` part is left alone; other slides may still
 * reference it.
 */
export const removeSlideNotes = (slide: SlideData): void => {
  const notesPartName = findNotesPartName(slide);
  if (notesPartName === null) return;
  const pkg = slide[INTERNAL_PACKAGE];
  pkg.removePart(notesPartName);
  pkg.removePart(relsPartNameFor(notesPartName));
  const slideRels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (slideRels === null) return;
  slideRels.items = slideRels.items.filter((r) => r.type !== REL_TYPES.notesSlide);
  pkg.setRels(slide[SLIDE_PART_NAME], slideRels);
};

/**
 * Removes the speaker-notes part from every slide that has one.
 * Built on `removeSlideNotes`. Returns the number of slides
 * stripped. Useful as a privacy/sharing helper before exporting a
 * deck whose notes contain internal commentary.
 */
export const clearAllSlideNotes = (pres: PresentationData): number => {
  let n = 0;
  for (const slide of getSlides(pres)) {
    if (findNotesPartName(slide) === null) continue;
    removeSlideNotes(slide);
    n++;
  }
  return n;
};

// ---------------------------------------------------------------------------
// Shape image replacement.

// ---------------------------------------------------------------------------
// Slide size.

/**
 * Width × height of the slide canvas, in EMU. `type` is PowerPoint's
 * aspect-ratio hint (`screen4x3`, `screen16x9`, ...); the actual size
 * is always `width` × `height`.
 */
export interface SlideSize {
  readonly width: Emu;
  readonly height: Emu;
  readonly type?: string;
}

/** Returns the slide canvas size, or `null` if `presentation.xml` omits it. */
export const getSlideSize = (pres: PresentationData): SlideSize | null => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (presPart === null) return null;
  const root = parseXml(decode(presPart.data)).root;
  const model = readPresentationPart(root);
  if (model.slideSize === null) return null;
  return {
    width: model.slideSize.cx as Emu,
    height: model.slideSize.cy as Emu,
    ...(model.slideSize.type !== undefined ? { type: model.slideSize.type } : {}),
  };
};

const NAME_SLD_SZ_FN = qname('p', 'sldSz', NS.pml);
const ATTR_CX = qname('', 'cx', '');
const ATTR_CY = qname('', 'cy', '');
const ATTR_TYPE = qname('', 'type', '');
const NAME_SLD_ID_LST_FN = qname('p', 'sldIdLst', NS.pml);

/**
 * Sets the slide canvas size. Creates `<p:sldSz>` when absent, replaces
 * its attributes when present. The `type` hint is preserved as given.
 *
 * Schema ordering: `<p:sldSz>` follows `<p:sldIdLst>` per ECMA-376
 * §19.2.1.26; we insert at the correct position when bootstrapping.
 */
export const setSlideSize = (pres: PresentationData, opts: SlideSize): void => {
  const pkg = pres[INTERNAL_PACKAGE];
  const presPart = pkg.getPart(PRES_PART_NAME);
  if (!presPart) throw new Error('presentation.xml is missing');
  const doc = parseXml(decode(presPart.data));

  let sldSz = firstChildElement(doc.root, NAME_SLD_SZ_FN);
  if (sldSz === null) {
    sldSz = elem(NAME_SLD_SZ_FN);
    const sldIdLst = firstChildElement(doc.root, NAME_SLD_ID_LST_FN);
    if (sldIdLst !== null) {
      const idx = doc.root.children.indexOf(sldIdLst);
      doc.root.children.splice(idx + 1, 0, sldSz);
    } else {
      doc.root.children.push(sldSz);
    }
  }

  sldSz.attrs = [attr(ATTR_CX, String(opts.width)), attr(ATTR_CY, String(opts.height))];
  if (opts.type !== undefined) sldSz.attrs.push(attr(ATTR_TYPE, opts.type));

  presPart.data = encode(serializeXml(doc));
};

import { emu as emuValue } from '../units.ts';

/** 10in × 7.5in (`screen4x3`). */
export const SLIDE_SIZE_4_3: SlideSize = {
  width: emuValue(9144000),
  height: emuValue(6858000),
  type: 'screen4x3',
};

/** 13.333in × 7.5in (`screen16x9`) — Office 2013+ default. */
export const SLIDE_SIZE_16_9: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(6858000),
  type: 'screen16x9',
};

/** 13.333in × 8.33in (`screen16x10`). */
export const SLIDE_SIZE_16_10: SlideSize = {
  width: emuValue(12192000),
  height: emuValue(7620000),
  type: 'screen16x10',
};
