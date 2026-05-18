// Shape image replacement.

import { getPictureEmbedRId } from '../../internal/drawingml/index.ts';
import {
  type ImageFormat,
  contentTypeForFormat,
  detectImageFormat,
  extensionForFormat,
  partName,
  resolveTarget,
} from '../../internal/opc/index.ts';
import {
  INTERNAL_PACKAGE,
  SHAPE_ELEMENT,
  SHAPE_SLIDE,
  SHAPE_SNAPSHOT,
  SLIDE_PART_NAME,
  type SlideShapeData,
} from '../_internal-symbols.ts';
// ---------------------------------------------------------------------------

/**
 * Replaces a picture's media with `bytes`. Same-format replacements
 * write in place; cross-format replacements allocate a new media part
 * and repoint the rel. The original geometry — crop, sizing, transform —
 * is preserved.
 */
export const setShapeImage = (
  shape: SlideShapeData,
  bytes: Uint8Array,
  options: { format?: ImageFormat } = {},
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
    return;
  }

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
};
