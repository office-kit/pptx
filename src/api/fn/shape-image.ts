// Shape image replacement.

import { getPictureEmbedRId } from '../../internal/drawingml/index.ts';
import { readPosition, readSize, setPosition, setSize } from '../../internal/drawingml/index.ts';
import {
  type ImageFormat,
  contentTypeForFormat,
  detectImageFormat,
  extensionForFormat,
  partName,
  readImagePixelSize,
  resolveTarget,
} from '../../internal/opc/index.ts';
import type { Emu } from '../units.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';
import { commitSlideData, refreshSlideData } from './_helpers.ts';
// ---------------------------------------------------------------------------

/**
 * How an image fills its target box:
 *
 *   - `'fill'` — stretch to the exact `w × h`, ignoring aspect ratio
 *     (the historical behavior, and the default for back-compat).
 *   - `'contain'` — scale to fit inside the box preserving aspect ratio,
 *     then center the result. The box's empty margins stay transparent.
 */
export type ImageFit = 'fill' | 'contain';

/**
 * Computes the placed rectangle for an image inside a `(x, y, w, h)` box.
 *
 * For `'contain'` the image is scaled by the smaller of the two axis
 * ratios so it fits entirely, then centered in the leftover space. When
 * the natural size is unknown (non-PNG/JPEG, or an unreadable header)
 * `naturalSize` is `null` and we fall back to `'fill'` — callers pass the
 * raw bytes; measuring is best-effort, not a hard requirement.
 */
export const fitImageRect = (
  box: { x: Emu; y: Emu; w: Emu; h: Emu },
  fit: ImageFit,
  naturalSize: { width: number; height: number } | null,
): { x: Emu; y: Emu; w: Emu; h: Emu } => {
  if (fit === 'fill' || naturalSize === null) return box;
  const scale = Math.min(box.w / naturalSize.width, box.h / naturalSize.height);
  const w = Math.round(naturalSize.width * scale);
  const h = Math.round(naturalSize.height * scale);
  const x = box.x + Math.round((box.w - w) / 2);
  const y = box.y + Math.round((box.h - h) / 2);
  return { x: x as Emu, y: y as Emu, w: w as Emu, h: h as Emu };
};

/**
 * Replaces a picture's media with `bytes`. Same-format replacements
 * write in place; cross-format replacements allocate a new media part
 * and repoint the rel. The geometry — crop, transform — is preserved.
 *
 * Pass `options.fit: 'contain'` to re-fit the picture's extent to the
 * replacement image's aspect ratio, inscribed and centered in the shape's
 * current `(off, ext)` box. The default `'fill'` leaves the extent as-is
 * (the historical behavior). When the new image's natural size can't be
 * measured (non-PNG/JPEG), `'contain'` is a no-op rather than an error.
 */
export const setShapeImage = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat; fit?: ImageFit } = {},
): void => {
  if (shape[SHAPE_SNAPSHOT].kind !== 'picture') {
    throw new Error(
      `setShapeImage only works on picture shapes; ${shape[SHAPE_SNAPSHOT].kind} is not one`,
    );
  }
  const format = options.format ?? detectImageFormat(bytes);
  if (format === null) {
    throw new Error(
      'setShapeImage: could not detect image format. Pass options.format explicitly.',
    );
  }
  const rEmbed = getPictureEmbedRId(shape[SHAPE_ELEMENT]);
  if (rEmbed === null) {
    throw new Error(`picture "${shape[SHAPE_SNAPSHOT].name}" has no r:embed`);
  }
  const slide = shape[SHAPE_SLIDE];
  const pkg = slide[INTERNAL_PACKAGE];
  const rels = pkg.getRels(slide[SLIDE_PART_NAME]);
  if (rels === null) throw new Error(`slide ${slide[SLIDE_PART_NAME]} has no rels`);
  const rel = rels.items.find((r) => r.id === rEmbed);
  if (!rel) throw new Error(`slide rels missing entry for r:embed="${rEmbed}"`);

  const mediaName = rel.target.startsWith('/')
    ? partName(rel.target)
    : resolveTarget(slide[SLIDE_PART_NAME], rel.target);
  const newExtension = extensionForFormat(format);
  const newContentType = contentTypeForFormat(format);
  const dotIdx = mediaName.lastIndexOf('.');
  const currentExtension = dotIdx >= 0 ? mediaName.slice(dotIdx + 1).toLowerCase() : '';

  if (currentExtension === newExtension) {
    const part = pkg.getPart(mediaName);
    if (!part) throw new Error(`media part missing: ${mediaName}`);
    part.data = bytes;
    part.contentType = newContentType;
  } else {
    let nextN = 1;
    const mediaPathRegex = /^\/ppt\/media\/image(\d+)\./;
    for (const p of pkg.parts) {
      const m = p.name.match(mediaPathRegex);
      if (m?.[1] !== undefined) {
        const num = Number.parseInt(m[1], 10);
        if (Number.isFinite(num) && num >= nextN) nextN = num + 1;
      }
    }
    const newPartName = partName(`/ppt/media/image${nextN}.${newExtension}`);
    const hasDefault = pkg.contentTypes.defaults.some(
      (d) => d.extension.toLowerCase() === newExtension,
    );
    if (!hasDefault) {
      pkg.contentTypes.defaults.push({ extension: newExtension, contentType: newContentType });
    }
    pkg.addPart(newPartName, newContentType, bytes);
    rel.target = `../media/image${nextN}.${newExtension}`;
    pkg.setRels(slide[SLIDE_PART_NAME], rels);
  }

  // 'contain' re-fits the picture's extent to the new image's aspect ratio
  // inside the shape's current box; 'fill' (the default) leaves geometry as
  // PowerPoint had it. Natural size is best-effort — unreadable headers
  // leave the box unchanged.
  if (options.fit === 'contain') {
    const pos = readPosition(shape[SHAPE_ELEMENT], 'picture');
    const size = readSize(shape[SHAPE_ELEMENT], 'picture');
    const natural = readImagePixelSize(bytes);
    if (pos !== null && size !== null && natural !== null) {
      const fitted = fitImageRect(
        { x: pos.x as Emu, y: pos.y as Emu, w: size.w as Emu, h: size.h as Emu },
        'contain',
        natural,
      );
      setPosition(shape[SHAPE_ELEMENT], 'picture', fitted.x, fitted.y);
      setSize(shape[SHAPE_ELEMENT], 'picture', fitted.w, fitted.h);
      commitSlideData(slide);
      refreshSlideData(slide);
    }
  }
};
