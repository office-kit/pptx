// Picture opacity and cropping.
import { getSlides } from './slide-query.ts';

import { getPictureEmbedRId } from '../../internal/drawingml/index.ts';
import {
  type ImageFormat,
  detectImageFormat,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import { type SlideLayoutType } from '../../internal/presentationml/index.ts';
import {
  NS,
  type XmlElement,
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
  SLIDE_SHAPES,
  type SlideData,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitAndRefresh } from './_helpers.ts';
import { getSlideLayoutPartName } from './layouts.ts';
import { getPresentationTheme, getSlideLayoutName, getSlideLayoutType } from './package.ts';
import {
  getShapeBounds,
  getShapeFlip,
  getSlideLayout,
  resolveDrawingColor,
  shapesOverlap,
} from './shapes.ts';
import { getSlideSize } from './features.ts';

// ---------------------------------------------------------------------------
// Picture opacity — `<a:alphaModFix>` inside the picture's `<a:blip>`.
//
// `amt` is ECMA-376's ST_PositiveFixedPercentage (0–100000, scale 1/1000
// of a percent). PowerPoint defaults to fully opaque when the element
// is absent. Pass `null` to remove a prior `<a:alphaModFix>`.

const NAME_ALPHA_MOD_FIX_FN = qname('a', 'alphaModFix', NS.dml);
const ATTR_AMT_FN = qname('', 'amt', '');

/**
 * Sets the picture's opacity (0–1 fraction; `1` is fully opaque, `0`
 * fully transparent). Pass `null` to remove an existing opacity
 * override and restore PowerPoint's default behavior.
 *
 * Throws for non-picture shapes and on opacities outside `[0, 1]`.
 */
/**
 * Returns the embedded image bytes for a picture shape, or `null`
 * when the shape isn't a picture or has no `r:embed` reference
 * (external images aren't followed).
 *
 * The returned `Uint8Array` is a live view into the package media
 * part — treat it as read-only; copy if you need an independent
 * buffer.
 */
export const getShapeImageBytes = (shape: SlideShapeData): Uint8Array | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
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
 * `true` when the shape's text body carries any visible characters.
 * Tighter than checking `getShapeText(shape) !== ''` because it
 * doesn't allocate the concatenated string.
 */
export const hasShapeText = (shape: SlideShapeData): boolean => {
  const text = shape[SHAPE_SNAPSHOT].text;
  return typeof text === 'string' && text.length > 0;
};

/**
 * `true` when the shape carries an embedded image — either a
 * `<p:pic>` picture or a `<p:spPr>/<a:blipFill>` image fill on a
 * regular shape. External `r:link` references count too.
 */
export const hasShapeImage = (shape: SlideShapeData): boolean => {
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    return getPictureEmbedRId(shape[SHAPE_ELEMENT]) !== null;
  }
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return false;
  return firstChildElement(spPr, qname('a', 'blipFill', NS.dml)) !== null;
};

/**
 * Dense per-slide image count array. Counts every shape that
 * `hasShapeImage` matches on each slide. Rounds out the density-array
 * family alongside the chart, table, shape, text, and comment
 * counters.
 */
export const getPresentationImageCountsBySlide = (pres: PresentationData): ReadonlyArray<number> =>
  getSlides(pres).map((s) => s[SLIDE_SHAPES].filter((sh) => hasShapeImage(sh)).length);

/**
 * Returns every shape on the slide that is mirrored — horizontally
 * (`flipH`), vertically (`flipV`), or both.
 */
export const findFlippedShapes = (slide: SlideData): ReadonlyArray<SlideShapeData> =>
  slide[SLIDE_SHAPES].filter((s) => {
    const flip = getShapeFlip(s);
    return flip !== null && (flip.horizontal || flip.vertical);
  });

/**
 * Returns every unordered pair of shapes on the slide whose
 * bounding boxes overlap. Built on `shapesOverlap`. Pairs are
 * returned with `a` strictly preceding `b` in document order, and
 * each pair appears at most once.
 *
 * Useful for layout audits — "do any boxes collide on this slide?"
 * Shapes without `<a:xfrm>` bounds never overlap anything.
 */
export const findOverlappingShapePairs = (
  slide: SlideData,
): ReadonlyArray<readonly [SlideShapeData, SlideShapeData]> => {
  const shapes = slide[SLIDE_SHAPES];
  const out: (readonly [SlideShapeData, SlideShapeData])[] = [];
  for (let i = 0; i < shapes.length; i++) {
    for (let j = i + 1; j < shapes.length; j++) {
      if (shapesOverlap(shapes[i]!, shapes[j]!)) {
        out.push([shapes[i]!, shapes[j]!] as const);
      }
    }
  }
  return out;
};

/**
 * Returns every shape on the slide whose bounding box extends past
 * the slide canvas (`getSlideSize(pres)`). Useful audit helper for
 * catching shapes that PowerPoint will silently render off-screen
 * or clip on export. Shapes without `<a:xfrm>` bounds are skipped.
 *
 * If the presentation has no slide-size declared, every positioned
 * shape is returned (caller can't audit against an absent canvas).
 */
export const findShapesOutsideCanvas = (
  slide: SlideData,
  pres: PresentationData,
): ReadonlyArray<SlideShapeData> => {
  const size = getSlideSize(pres);
  const out: SlideShapeData[] = [];
  for (const shape of slide[SLIDE_SHAPES]) {
    const b = getShapeBounds(shape);
    if (b === null) continue;
    if (size === null) {
      out.push(shape);
      continue;
    }
    if (b.x < 0 || b.y < 0 || b.x + b.w > size.width || b.y + b.h > size.height) {
      out.push(shape);
    }
  }
  return out;
};

/**
 * Every slide whose layout's `cSld@name` matches the given string.
 * Useful for batch operations on slides sharing a layout — for
 * example, restyling every "Title and Content" slide in a deck.
 *
 * Matching is exact (case-sensitive). Slides without a resolved
 * layout are skipped.
 */
export const findSlidesByLayoutName = (
  pres: PresentationData,
  layoutName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutName(layout) === layoutName) out.push(slide);
  }
  return out;
};

/**
 * Every slide whose resolved layout part name equals `layoutPartName`
 * (e.g. `'/ppt/slideLayouts/slideLayout3.xml'`). Stable across
 * template-name collisions and locale renames — keyed on the actual
 * package path. Pair to `findSlidesByLayoutName` /
 * `findSlidesByLayoutType` for cases where the caller already has a
 * layout part name in hand (typical when iterating over a layouts
 * collection programmatically).
 */
export const findSlidesByLayoutPartName = (
  pres: PresentationData,
  layoutPartName: string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutPartName(layout) === layoutPartName) out.push(slide);
  }
  return out;
};

/**
 * Every slide whose layout `@type` (e.g. `'title'`, `'blank'`,
 * `'obj'`) matches. Sibling of `findSlidesByLayoutName`, but keyed
 * on the OOXML layout-type enum rather than the human-facing name —
 * stable across locales and template providers.
 */
export const findSlidesByLayoutType = (
  pres: PresentationData,
  layoutType: SlideLayoutType | string,
): ReadonlyArray<SlideData> => {
  const out: SlideData[] = [];
  for (const slide of getSlides(pres)) {
    const layout = getSlideLayout(slide);
    if (layout !== null && getSlideLayoutType(layout) === layoutType) out.push(slide);
  }
  return out;
};

/**
 * Returns the package part name (`/ppt/media/imageN.ext`) of
 * whichever image the shape carries — picture (`<p:pic>`) or
 * image-fill (`<a:blipFill>` nested under `<p:spPr>`). Returns
 * `null` when the shape has no embedded image, or the rel points
 * at an external `r:link` rather than an internal target.
 *
 * Useful for addressing the media part directly with
 * `setMediaPartBytes` or `readPackagePart`.
 */
/**
 * Returns the external URL of the picture when its `<a:blip>` carries an
 * `r:link` (external) relationship rather than an `r:embed`. Returns
 * `null` for embedded pictures, non-picture shapes, or when the
 * relationship doesn't resolve.
 *
 * PowerPoint emits `r:link` when the user inserts via "Link to file"
 * instead of "Insert Picture". The bytes live outside the package, so
 * `getShapeImageBytes` can't render them — readers / preview tools
 * should fall back to this URL or a placeholder.
 */
/**
 * Returns `true` when the picture's `<a:blip>` carries `<a:grayscl/>`
 * — PowerPoint's "Color > Grayscale" recolor preset. Renderers can
 * project this onto a CSS `filter: grayscale(100%)` or an SVG
 * `<feColorMatrix>` desaturation.
 */
export const isShapeImageGrayscale = (shape: SlideShapeData): boolean => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  return blip !== null && firstChildElement(blip, qname('a', 'grayscl', NS.dml)) !== null;
};

/**
 * Returns the threshold of the picture's `<a:blip><a:biLevel thresh="…"/>`
 * effect — PowerPoint's "Color > Black and White" preset. Threshold is
 * a percent (0..100); pixels brighter become white, darker become black.
 * Returns `null` when no biLevel transform is set.
 */
export const getShapeImageBiLevelThreshold = (shape: SlideShapeData): number | null => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const biLevel = firstChildElement(blip, qname('a', 'biLevel', NS.dml));
  if (!biLevel) return null;
  const t = getAttrValue(biLevel, qname('', 'thresh', ''));
  if (t === null) return null;
  let n = Number.parseFloat(t);
  if (!Number.isFinite(n)) return null;
  if (Math.abs(n) > 1) n = n / 100000;
  return n * 100;
};

/**
 * Reads the picture's duotone color transform from `<a:blip><a:duotone>`.
 * PowerPoint emits two `<a:srgbClr>` (or scheme color) children for a
 * two-color duotone effect — typical "Picture Tools › Recolor".
 * Returns `null` when no duotone is set.
 */
export const getShapeImageDuotone = (
  pres: PresentationData,
  shape: SlideShapeData,
): { firstColor: string | null; secondColor: string | null } | null => {
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const duotone = firstChildElement(blip, qname('a', 'duotone', NS.dml));
  if (!duotone) return null;
  const theme = getPresentationTheme(pres);
  const colors: Array<string | null> = [];
  for (const c of duotone.children) {
    if (c.kind !== 'element' || c.name.namespaceURI !== NS.dml) continue;
    if (
      c.name.localName === 'srgbClr' ||
      c.name.localName === 'schemeClr' ||
      c.name.localName === 'sysClr' ||
      c.name.localName === 'prstClr'
    ) {
      colors.push(resolveDrawingColor(c, theme));
      if (colors.length === 2) break;
    }
  }
  return {
    firstColor: colors[0] ?? null,
    secondColor: colors[1] ?? null,
  };
};

export const getShapeImageLinkUrl = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;
  // Find the blip element on either picture or shape-with-image-fill.
  let blip: XmlElement | null = null;
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
    if (blipFill) blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  } else {
    const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
    if (spPr) {
      const bf = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
      if (bf) blip = firstChildElement(bf, qname('a', 'blip', NS.dml));
    }
  }
  if (!blip) return null;
  const rLink = getAttrValue(blip, qname('r', 'link', NS.officeDocRels));
  if (!rLink) return null;
  const rel = rels.items.find((r) => r.id === rLink);
  if (!rel || rel.targetMode !== 'External') return null;
  return rel.target;
};

export const getShapeImagePartName = (shape: SlideShapeData): string | null => {
  const slide = shape[SHAPE_SLIDE];
  const rels = slide[INTERNAL_PACKAGE].getRels(slide[SLIDE_PART_NAME]);
  if (!rels) return null;

  const resolve = (rEmbed: string | null): string | null => {
    if (rEmbed === null) return null;
    const rel = rels.items.find((r) => r.id === rEmbed);
    if (!rel || rel.targetMode === 'External') return null;
    const name = rel.target.startsWith('/')
      ? partName(rel.target)
      : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
    return name;
  };

  // Picture shape: <p:pic><p:blipFill><a:blip r:embed="..."/>.
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
    return resolve(rEmbed);
  }

  // Other shapes with image fill: <p:spPr><a:blipFill><a:blip r:embed="..."/>.
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  return resolve(getAttrValue(blip, qname('r', 'embed', NS.officeDocRels)));
};

/**
 * Returns the bytes of the image used as this shape's *fill*
 * (`<a:blipFill>` nested under `<p:spPr>`, as written by
 * `setShapeImageFill`). Distinct from `getShapeImageBytes`, which only
 * applies to `<p:pic>` picture shapes.
 *
 * Returns null if the shape has no image fill, the blip has no
 * `r:embed`, or the embed points at an external `r:link`.
 */
export const getShapeImageFillBytes = (shape: SlideShapeData): Uint8Array | null => {
  const spPr = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'spPr', NS.pml));
  if (!spPr) return null;
  const blipFill = firstChildElement(spPr, qname('a', 'blipFill', NS.dml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const rEmbed = getAttrValue(blip, qname('r', 'embed', NS.officeDocRels));
  if (rEmbed === null) return null;
  const slide = shape[SHAPE_SLIDE];
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
 * Returns the image format token (`'png'`, `'jpeg'`, …) for whichever
 * image bytes the shape carries — picture (`<p:pic>`) or image-fill
 * (`<a:blipFill>` on `<p:spPr>`). Returns `null` if the shape has no
 * embedded image or the bytes don't match a recognized signature.
 */
export const getShapeImageFormat = (shape: SlideShapeData): ImageFormat | null => {
  const bytes = getShapeImageBytes(shape) ?? getShapeImageFillBytes(shape);
  if (bytes === null) return null;
  return detectImageFormat(bytes);
};

/**
 * Reads the picture's opacity (0–1 fraction). Returns `null` when no
 * `<a:alphaModFix>` is present (PowerPoint treats absence as fully
 * opaque); returns `1` when an explicit alphaModFix sets full opacity.
 */
export const getShapeImageOpacity = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const alpha = firstChildElement(blip, qname('a', 'alphaModFix', NS.dml));
  if (!alpha) return null;
  const amt = getAttrValue(alpha, qname('', 'amt', ''));
  if (amt === null) return 1;
  const n = Number.parseInt(amt, 10);
  if (!Number.isFinite(n)) return null;
  return n / 100000;
};

/**
 * Reads the picture's crop fractions. Returns `null` when no
 * `<a:srcRect>` is present; otherwise returns a fully-populated object
 * with every side filled in (0 for omitted sides on disk).
 */
export const getShapeImageCrop = (shape: SlideShapeData): ImageCrop | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const srcRect = firstChildElement(blipFill, qname('a', 'srcRect', NS.dml));
  if (!srcRect) return null;
  const parseSide = (local: string): number => {
    const v = getAttrValue(srcRect, qname('', local, ''));
    if (v === null) return 0;
    const n = Number.parseInt(v, 10);
    return Number.isFinite(n) ? n / 100000 : 0;
  };
  return {
    left: parseSide('l'),
    top: parseSide('t'),
    right: parseSide('r'),
    bottom: parseSide('b'),
  };
};

/**
 * Adjusts the picture's brightness by writing `<a:lumOff val="…"/>`
 * inside `<a:blip>`. The value is a -1..1 fraction:
 *
 *   - `1`     → +100% brightness
 *   - `0` or `null` → no offset (any prior `<a:lumOff>` is removed)
 *   - `-1`    → -100% brightness
 *
 * Throws for non-picture shapes and on values outside [-1, 1].
 *
 * Note: PowerPoint's "Picture Format › Corrections" UI couples this
 * with `<a:lumMod>` for some presets; this primitive sets only
 * `lumOff` to keep the surface honest. Read it back via
 * `getShapeImageBrightness`.
 */
export const setShapeImageBrightness = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageBrightness only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumOff'),
  );
  if (value !== null && value !== 0) {
    if (!Number.isFinite(value) || value < -1 || value > 1) {
      throw new RangeError(`brightness must be in [-1, 1], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumOff', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Adjusts the picture's contrast by writing `<a:lumMod val="…"/>` on
 * `<a:blip>`. The value is a 0..2 fraction:
 *
 *   - `1` or `null` → no modulation (any prior `<a:lumMod>` is removed)
 *   - `0.5`         → 50% of original luminance variance (washed out)
 *   - `1.5`         → 150% (boosted contrast; PowerPoint clamps to
 *                       what the renderer supports)
 *
 * Throws on non-picture shapes and on values outside `[0, 2]`. The
 * primitive maps directly to `ST_PositiveFixedPercentage` × 100000.
 */
export const setShapeImageContrast = (shape: SlideShapeData, value: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageContrast only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');
  blip.children = blip.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'lumMod'),
  );
  if (value !== null && value !== 1) {
    if (!Number.isFinite(value) || value < 0 || value > 2) {
      throw new RangeError(`contrast must be in [0, 2], got ${value}`);
    }
    blip.children.push(
      elem(qname('a', 'lumMod', NS.dml), {
        attrs: [attr(qname('', 'val', ''), String(Math.round(value * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

/**
 * Reads the picture's contrast modulation (the `<a:lumMod>` fraction
 * in [0, 2]). Returns `null` when no `<a:lumMod>` is present.
 */
export const getShapeImageContrast = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumMod = firstChildElement(blip, qname('a', 'lumMod', NS.dml));
  if (!lumMod) return null;
  const v = getAttrValue(lumMod, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

/**
 * Reads the picture's brightness offset (the `<a:lumOff>` fraction
 * in [-1, 1]). Returns `null` when no `<a:lumOff>` is present.
 */
export const getShapeImageBrightness = (shape: SlideShapeData): number | null => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') return null;
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) return null;
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) return null;
  const lumOff = firstChildElement(blip, qname('a', 'lumOff', NS.dml));
  if (!lumOff) return null;
  const v = getAttrValue(lumOff, qname('', 'val', ''));
  if (v === null) return null;
  const n = Number.parseInt(v, 10);
  return Number.isFinite(n) ? n / 100000 : null;
};

export const setShapeImageOpacity = (shape: SlideShapeData, opacity: number | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageOpacity only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const blipFill = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'blipFill', NS.pml));
  if (!blipFill) throw new Error('picture has no <p:blipFill>');
  const blip = firstChildElement(blipFill, qname('a', 'blip', NS.dml));
  if (!blip) throw new Error('picture <p:blipFill> has no <a:blip>');

  blip.children = blip.children.filter(
    (c) =>
      !(
        c.kind === 'element' &&
        c.name.namespaceURI === NS.dml &&
        c.name.localName === 'alphaModFix'
      ),
  );

  if (opacity !== null) {
    if (!Number.isFinite(opacity) || opacity < 0 || opacity > 1) {
      throw new RangeError(`opacity must be in [0, 1], got ${opacity}`);
    }
    blip.children.push(
      elem(NAME_ALPHA_MOD_FIX_FN, {
        attrs: [attr(ATTR_AMT_FN, String(Math.round(opacity * 100000)))],
      }),
    );
  }
  commitAndRefresh(shape);
};

// ---------------------------------------------------------------------------
// Picture cropping — `<a:srcRect>` inside the picture's `<p:blipFill>`.
//
// Percentages are 0-1 fractions per side, converted to ECMA-376's
// `ST_Percentage` units (1/1000 of a percent, so 0.25 → "25000"). Pass
// `null` to remove an existing crop.

/** Crop a picture by fraction of each side. Omitted sides default to 0. */
export interface ImageCrop {
  readonly left?: number;
  readonly top?: number;
  readonly right?: number;
  readonly bottom?: number;
}

const NAME_BLIP_FILL_FN = qname('p', 'blipFill', NS.pml);
const NAME_SRC_RECT_FN = qname('a', 'srcRect', NS.dml);
const NAME_BLIP_FN = qname('a', 'blip', NS.dml);
const ATTR_CROP_L = qname('', 'l', '');
const ATTR_CROP_T = qname('', 't', '');
const ATTR_CROP_R = qname('', 'r', '');
const ATTR_CROP_B = qname('', 'b', '');

const fractionToST = (n: number | undefined): string | null => {
  if (n === undefined || n === 0) return null;
  if (!Number.isFinite(n) || n < 0 || n >= 1) {
    throw new RangeError(`crop fraction must be in [0, 1), got ${n}`);
  }
  return String(Math.round(n * 100000));
};

/**
 * Sets (or clears) a `<a:srcRect>` on a picture shape, cropping the
 * embedded image by the given fraction on each side. Pass `null` to
 * remove an existing crop.
 *
 * Fractions are in `[0, 1)` per side. `{ left: 0.25 }` clips 25% off
 * the left edge; the visible image stretches to fill the original
 * frame. The shape's geometry (`<a:xfrm>`) is unchanged.
 */
export const setShapeImageCrop = (shape: SlideShapeData, crop: ImageCrop | null): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImageCrop only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const pic = shape[SHAPE_ELEMENT];
  const blipFill = firstChildElement(pic, NAME_BLIP_FILL_FN);
  if (!blipFill) throw new Error('picture has no <p:blipFill>');

  // Remove any existing srcRect first.
  blipFill.children = blipFill.children.filter(
    (c) =>
      !(c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'srcRect'),
  );

  if (crop === null) {
    commitAndRefresh(shape);
    return;
  }

  const attrs: Array<ReturnType<typeof attr>> = [];
  const l = fractionToST(crop.left);
  const t = fractionToST(crop.top);
  const r = fractionToST(crop.right);
  const b = fractionToST(crop.bottom);
  if (l !== null) attrs.push(attr(ATTR_CROP_L, l));
  if (t !== null) attrs.push(attr(ATTR_CROP_T, t));
  if (r !== null) attrs.push(attr(ATTR_CROP_R, r));
  if (b !== null) attrs.push(attr(ATTR_CROP_B, b));

  // <a:srcRect> sits between <a:blip> and <a:stretch> per the schema.
  const srcRect = elem(NAME_SRC_RECT_FN, { attrs });
  const blipIdx = blipFill.children.findIndex(
    (c) => c.kind === 'element' && c.name.namespaceURI === NS.dml && c.name.localName === 'blip',
  );
  if (blipIdx === -1) {
    // No <a:blip>? Just prepend the srcRect.
    blipFill.children.unshift(srcRect);
  } else {
    blipFill.children.splice(blipIdx + 1, 0, srcRect);
  }
  commitAndRefresh(shape);
};

void NAME_BLIP_FN;
