// Slide-level background.

import {
  type GradientFillOptions,
  readFlip,
  readPosition,
  readRotation,
  readSize,
  setSolidFill,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  type ImageFormat,
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  nextRelId,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { REL_TYPES, readShapeTreeFromCsldRoot } from '../../internal/presentationml/index.ts';
import type { OpcPackage } from '../../internal/parts/index.ts';
import type { PartName } from '../../internal/opc/index.ts';
import {
  NS,
  type XmlElement,
  attr,
  elem,
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
  SHAPE_SNAPSHOT,
  SLIDE_DOCUMENT,
  SLIDE_PART,
  SLIDE_PART_NAME,
  SLIDE_SHAPES,
  type SlideData,
  type SlideLayoutData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { NAME_CSLD, commitSlideData, decode, refreshSlideData, setOpcDefault } from './_helpers.ts';
import { getPresentationTheme } from './theme.ts';
import {
  NAME_A_GRAD_FILL,
  NAME_A_GS_LST,
  NAME_A_LIN,
  type ShapeBounds,
  readColorFromContainer,
  resolveDrawingColor,
} from './shapes.ts';

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

// Wraps a layout / master part as a read-only `SlideData` so its decorative
// shapes can be read (and rendered) through the very same `getShape*` helpers
// as slide shapes — crucially, picture relationships (logos) resolve against
// THIS part's rels. The handle is for reading only; mutating helpers would
// commit against a part this wrapper doesn't track.
const partDecorationShapes = (
  pkg: OpcPackage,
  thePartName: PartName,
  csldType: 'sldLayout' | 'sldMaster',
): ReadonlyArray<SlideShapeData> => {
  const part = pkg.getPart(thePartName);
  if (part === null) return [];
  const doc = parseXml(decode(part.data));
  const { shapes } = readShapeTreeFromCsldRoot(doc.root, csldType);
  const data: SlideData = {
    [INTERNAL_PACKAGE]: pkg,
    [SLIDE_PART_NAME]: thePartName,
    [SLIDE_DOCUMENT]: doc,
    [SLIDE_PART]: { shapes, root: doc.root },
    [SLIDE_SHAPES]: [],
  };
  // Decoration only: placeholders are rendered through the slide's own
  // placeholders (which already cascade through the layout/master).
  data[SLIDE_SHAPES] = shapes
    .filter((s) => s.placeholderType === null && s.placeholderIdx === null)
    .map((snap) => ({
      [SHAPE_SLIDE]: data,
      [SHAPE_ELEMENT]: snap.element,
      [SHAPE_SNAPSHOT]: snap,
    }));
  return data[SLIDE_SHAPES];
};

/**
 * Non-placeholder decorative shapes on a slide layout — corner bars, divider
 * lines, **logos**, watermark text — as full `SlideShapeData` bound to the
 * layout part. Unlike `getSlideLayoutBackgroundShapes` (a flat summary that
 * drops pictures and groups), these work with every `getShape*` reader, so a
 * renderer can paint them with real geometry, fills, text and picture bytes.
 *
 * For reading / rendering only — the handles are bound to the layout part, not
 * a slide, so the mutating `setShape*` helpers should not be used on them.
 */
export const getSlideLayoutShapes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideShapeData> =>
  partDecorationShapes(pres[INTERNAL_PACKAGE], layout[LAYOUT_PART_NAME], 'sldLayout');

/**
 * Non-placeholder decorative shapes on the layout's slide master — the
 * template decoration shared by every layout (background logos, divider lines,
 * etc.). Same render-ready `SlideShapeData` view as `getSlideLayoutShapes`.
 */
export const getSlideMasterShapes = (
  pres: PresentationData,
  layout: SlideLayoutData,
): ReadonlyArray<SlideShapeData> => {
  const pkg = pres[INTERNAL_PACKAGE];
  const layoutName = layout[LAYOUT_PART_NAME];
  const rels = pkg.getRels(layoutName);
  const masterRel = rels?.items.find((r) => r.type === REL_TYPES.slideMaster);
  if (!masterRel) return [];
  const masterName = masterRel.target.startsWith('/')
    ? partName(masterRel.target)
    : resolveTarget(layoutName, masterRel.target);
  return partDecorationShapes(pkg, masterName, 'sldMaster');
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
