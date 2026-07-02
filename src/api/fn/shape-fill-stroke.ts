// Shape mutation: geometry, fill, stroke.

import { resolveDrawingColor } from './shape-color.ts';
import {
  type ArrowOptions,
  type GradientFillOptions,
  type LineDash,
  type PatternFillOptions,
  type StrokeOptions,
  clearFill as clearFillImpl,
  clearStroke as clearStrokeImpl,
  setAdjustValues as writeAdjustValues,
  setFlip as writeFlip,
  setGradientFill,
  setPatternFill,
  setNoFill as setNoFillImpl,
  setNoStroke as setNoStrokeImpl,
  setPosition as writePosition,
  setRotation as writeRotation,
  setSize as writeSize,
  setSolidFill,
  setSolidStroke,
  setStrokeArrow,
  setStrokeCap,
  setStrokeCompound,
  setStrokeJoin,
  setStrokeDash,
} from '../../internal/drawingml/index.ts';
import type { Emu } from '../units.ts';
import {
  contentTypeForFormat,
  detectImageFormat,
  emptyRels,
  extensionForFormat,
  type ImageFormat,
  nextRelId,
  partName,
} from '../../internal/opc/index.ts';
import { REL_TYPES } from '../../internal/presentationml/index.ts';
import {
  NS,
  attr,
  elem,
  firstChildElement,
  getAttrValue,
  qname,
} from '../../internal/xml/index.ts';
import {
  INTERNAL_PACKAGE,
  type PresentationData,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh, requireSpPr, setOpcDefault } from './_helpers.ts';
import { getPresentationTheme } from './theme.ts';
// ---------------------------------------------------------------------------
// Shape mutation — geometry.

/** Sets the shape's position in EMU. Companion to `setShapeSize`. */
export const setShapePosition = (shape: SlideShapeData, x: Emu, y: Emu): void => {
  writePosition(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, x, y);
  commitAndRefresh(shape);
};

/** Sets the shape's size in EMU. */
export const setShapeSize = (shape: SlideShapeData, w: Emu, h: Emu): void => {
  writeSize(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, w, h);
  commitAndRefresh(shape);
};

/**
 * Sets the shape's rotation in degrees (positive clockwise). Values are
 * normalized into `[0, 360)`; pass `0` to clear an existing rotation.
 */
export const setShapeRotation = (shape: SlideShapeData, degrees: number): void => {
  writeRotation(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, degrees);
  commitAndRefresh(shape);
};

/**
 * Sets the shape's preset-geometry adjust values (`<a:prstGeom><a:avLst>`),
 * replacing any guides already authored. Keys are ECMA-376 guide names and
 * values are the raw guide numbers (rounded to integers). The companion
 * reader is {@link getShapeAdjustValues}.
 *
 * The common use is the corner radius of the `roundRect` preset, whose `adj`
 * guide runs `0..50000` (thousandths of a percent of the shorter side;
 * `16667` ≈ the PowerPoint default, `0` = square corners, `50000` = fully
 * rounded). For example, `setShapeAdjustValues(shape, { adj: 5000 })` gives a
 * subtle 5% rounding.
 *
 * Throws when the shape has no preset geometry (`<a:prstGeom>`) — a custom or
 * inherited geometry has no adjust list to author.
 */
export const setShapeAdjustValues = (
  shape: SlideShapeData,
  values: Readonly<Record<string, number>>,
): void => {
  if (!writeAdjustValues(shape[SHAPE_ELEMENT], values)) {
    throw new Error('setShapeAdjustValues: shape has no preset geometry (<a:prstGeom>)');
  }
  commitAndRefresh(shape);
};

/** Sets the shape's flip flags. Properties default to current state when omitted. */
export const setShapeFlip = (
  shape: SlideShapeData,
  options: { horizontal?: boolean; vertical?: boolean },
): void => {
  writeFlip(shape[SHAPE_ELEMENT], shape[SHAPE_SNAPSHOT].kind, options);
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Shape mutation — fill / stroke.

/** Sets a solid fill on the shape (color in `#RRGGBB` or scheme token). */
export const setShapeFill = (shape: SlideShapeData, color: string): void => {
  setSolidFill(requireSpPr(shape), color);
  commitAndRefresh(shape);
};

/**
 * Sets a linear gradient fill on the shape. Stops must lie in `[0, 1]`;
 * `angleDeg` defaults to `90` (top → bottom).
 *
 * Example: red → blue top-to-bottom:
 *
 *   setShapeGradientFill(shape, {
 *     stops: [{ offset: 0, color: '#FF0000' }, { offset: 1, color: '#0000FF' }],
 *     angleDeg: 90,
 *   });
 */
export const setShapeGradientFill = (shape: SlideShapeData, options: GradientFillOptions): void => {
  setGradientFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Sets a preset pattern fill on the shape (e.g. `pct50`, `dkUpDiag`).
 *
 * `foreground` is the pattern stroke color; `background` fills behind
 * the pattern. Both accept `#RRGGBB`, bare `RRGGBB`, or scheme tokens
 * (`accent1`, `bg1`, ...).
 */
export const setShapePatternFill = (shape: SlideShapeData, options: PatternFillOptions): void => {
  setPatternFill(requireSpPr(shape), options);
  commitAndRefresh(shape);
};

/**
 * Reads back the pattern fill on a shape: returns the preset token
 * plus the foreground / background colors resolved against the theme.
 * Returns `null` when the shape has no `<a:pattFill>`.
 *
 * The preset string is the literal `ST_PresetPatternVal` token from
 * §20.1.10.49 — e.g. `'pct50'`, `'dkUpDiag'`, `'cross'`, `'wave'`.
 * Renderers can map it onto an SVG `<pattern>` definition.
 */
export const getShapePatternFill = (
  pres: PresentationData,
  shape: SlideShapeData,
): { preset: string; foreground: string; background: string } | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const pattFill = firstChildElement(spPr, qname('a', 'pattFill', NS.dml));
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
 * Sets a picture fill on the shape, embedding `bytes` as a new media
 * part and replacing any prior fill choice on the shape's `<p:spPr>`.
 *
 * The image stretches to fill the shape (`<a:stretch><a:fillRect/>`).
 * Format is detected from magic bytes; pass `options.format` to
 * override (useful for SVG or unusual extensions).
 *
 * Throws if the format can't be detected and isn't provided explicitly,
 * or if the shape kind doesn't carry a `<p:spPr>` (e.g. groups).
 */
export const setShapeImageFill = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
): void => {
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImageFill: could not detect image format. Pass options.format explicitly.',
    );
  }
  const contentType = contentTypeForFormat(format);
  const extension = extensionForFormat(format);
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];

  // Allocate /ppt/media/imageN.<ext> (shared with addSlideImage's
  // numbering — both feed off the same /ppt/media space).
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

  // Replace the shape's fill choice with <a:blipFill>.
  const spPr = requireSpPr(shape);
  const FILL_CHOICES = new Set([
    'noFill',
    'solidFill',
    'gradFill',
    'blipFill',
    'pattFill',
    'grpFill',
  ]);
  spPr.children = spPr.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        FILL_CHOICES.has(c.name.localName)
      ),
  );
  const blipName = qname('a', 'blip', NS.dml);
  const stretchName = qname('a', 'stretch', NS.dml);
  const fillRectName = qname('a', 'fillRect', NS.dml);
  const blipFillName = qname('a', 'blipFill', NS.dml);
  const blip = elem(blipName, { attrs: [attr(qname('r', 'embed', NS.officeDocRels), newRId)] });
  const stretch = elem(stretchName, { children: [elem(fillRectName)] });
  const blipFill = elem(blipFillName, { children: [blip, stretch] });
  // <a:blipFill> takes the same slot as <a:solidFill>; insert at the
  // current insertion index. We use the same heuristic as setSolidFill —
  // before <a:ln> / effectLst / scene3d / extLst.
  let insertAt = spPr.children.length;
  for (let i = 0; i < spPr.children.length; i++) {
    const c = spPr.children[i];
    if (c?.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (
      c.name.localName === 'ln' ||
      c.name.localName === 'effectLst' ||
      c.name.localName === 'effectDag' ||
      c.name.localName === 'scene3d' ||
      c.name.localName === 'sp3d' ||
      c.name.localName === 'extLst'
    ) {
      insertAt = i;
      break;
    }
  }
  spPr.children.splice(insertAt, 0, blipFill);
  commitAndRefresh(shape);
};

/** Sets `<a:noFill>` on the shape, leaving it transparent. */
export const setShapeNoFill = (shape: SlideShapeData): void => {
  setNoFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/**
 * Removes any fill choice from the shape; it then inherits its fill
 * from the layout / master placeholder it descends from.
 */
export const clearShapeFill = (shape: SlideShapeData): void => {
  clearFillImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Sets a solid-color outline on the shape. */
export const setShapeStroke = (
  shape: SlideShapeData,
  options: { color?: string; widthEmu?: number },
): void => {
  setSolidStroke(requireSpPr(shape), options as StrokeOptions);
  commitAndRefresh(shape);
};

/** Sets an explicit "no outline" on the shape. */
export const setShapeNoStroke = (shape: SlideShapeData): void => {
  setNoStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/** Reads back the shape's stroke dash style, or `null` if none. */
export const getShapeStrokeDash = (shape: SlideShapeData): LineDash | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const prstDash = firstChildElement(ln, qname('a', 'prstDash', NS.dml));
  if (!prstDash) return null;
  const v = getAttrValue(prstDash, qname('', 'val', ''));
  return (v as LineDash | null) ?? null;
};

/**
 * Reads back the shape's arrowhead on one end of `<a:ln>`, or `null`
 * when no `<a:headEnd>` / `<a:tailEnd>` is present.
 */
export const getShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
): ArrowOptions | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const ln = firstChildElement(spPr, qname('a', 'ln', NS.dml));
  if (!ln) return null;
  const arr = firstChildElement(ln, qname('a', end === 'head' ? 'headEnd' : 'tailEnd', NS.dml));
  if (!arr) return null;
  const type = getAttrValue(arr, qname('', 'type', ''));
  if (!type) return null;
  const width = getAttrValue(arr, qname('', 'w', ''));
  const length = getAttrValue(arr, qname('', 'len', ''));
  const result: {
    type: ArrowOptions['type'];
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  } = {
    type: type as ArrowOptions['type'],
  };
  if (width === 'sm' || width === 'med' || width === 'lg') result.width = width;
  if (length === 'sm' || length === 'med' || length === 'lg') result.length = length;
  return result;
};

/**
 * Sets the dash pattern for the shape's outline (`<a:prstDash>`). One
 * of ECMA-376's `ST_PresetLineDashVal` tokens:
 *
 *   `'solid'` | `'dot'` | `'dash'` | `'lgDash'` | `'dashDot'` |
 *   `'lgDashDot'` | `'lgDashDotDot'` | `'sysDash'` | `'sysDot'` |
 *   `'sysDashDot'` | `'sysDashDotDot'`
 *
 * Creates `<a:ln>` if absent. Pairs naturally with `setShapeStroke`:
 * users typically set a color + width first, then the dash.
 */
export const setShapeStrokeDash = (shape: SlideShapeData, dash: LineDash): void => {
  setStrokeDash(requireSpPr(shape), dash);
  commitAndRefresh(shape);
};

/**
 * Sets an arrowhead on one end of the shape's outline.
 *
 *   - `end: 'head'` writes `<a:headEnd>` (the start of the line).
 *   - `end: 'tail'` writes `<a:tailEnd>` (the end).
 *
 * Useful primarily on connector shapes added via `addSlideLine`.
 * `type: 'none'` clears the arrowhead.
 */
export const setShapeStrokeArrow = (
  shape: SlideShapeData,
  end: 'head' | 'tail',
  options: ArrowOptions,
): void => {
  setStrokeArrow(requireSpPr(shape), end, options);
  commitAndRefresh(shape);
};

/** Removes any outline override; the shape then inherits stroke from layout. */
export const clearShapeStroke = (shape: SlideShapeData): void => {
  clearStrokeImpl(requireSpPr(shape));
  commitAndRefresh(shape);
};

/**
 * Sets the line-cap style on the shape's outline (`<a:ln cap="…">`).
 *
 *   - `'rnd'`  — rounded ends.
 *   - `'sq'`   — square ends extending past the endpoint.
 *   - `'flat'` — square ends flush at the endpoint (the OOXML default).
 *
 * Pass `null` to clear the attribute so the cap inherits the default.
 * Creates `<a:ln>` if absent.
 */
export const setShapeStrokeCap = (
  shape: SlideShapeData,
  cap: 'rnd' | 'sq' | 'flat' | null,
): void => {
  setStrokeCap(requireSpPr(shape), cap);
  commitAndRefresh(shape);
};

/**
 * Sets the line-join style on the shape's outline. Picks one of the
 * three child-element variants of `<a:ln>`:
 *
 *   - `'round'` → `<a:round/>`
 *   - `'bevel'` → `<a:bevel/>`
 *   - `'miter'` → `<a:miter/>`
 *
 * Pass `null` to clear any prior join child so the shape inherits the
 * default. Creates `<a:ln>` if absent.
 */
export const setShapeStrokeJoin = (
  shape: SlideShapeData,
  join: 'round' | 'bevel' | 'miter' | null,
): void => {
  setStrokeJoin(requireSpPr(shape), join);
  commitAndRefresh(shape);
};

/**
 * Sets the compound-line style on the shape's outline
 * (`<a:ln cmpd="…">`) — single, double, triple, or thick/thin variants.
 * ECMA-376 §20.1.10.31 `ST_CompoundLine`. Pass `null` to clear the
 * attribute. Creates `<a:ln>` if absent.
 */
export const setShapeStrokeCompound = (
  shape: SlideShapeData,
  cmpd: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null,
): void => {
  setStrokeCompound(requireSpPr(shape), cmpd);
  commitAndRefresh(shape);
};
