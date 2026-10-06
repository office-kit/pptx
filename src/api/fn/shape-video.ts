import { NS } from '../../internal/xml/index.ts';
import type { SlideShapeData } from '../_internal-symbols.ts';
import { commitAndRefresh, requireSpPr } from './_helpers.ts';
import { getShapeMedia } from './media.ts';
import { setShapePreset } from './shape-fill-stroke.ts';
import { resetShapeImageColorEffects } from './shape-image-effects.ts';

// PowerPoint's Video Format > Reset clears the shape-level style while leaving
// the media picture fill (poster), crop, and transform intact; geometry is
// restored to the rectangle preset.
const VIDEO_FORMAT_CHILDREN = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
]);

/**
 * Restores a video to PowerPoint's default shape formatting.
 *
 * This is deliberately restricted to video media pictures. It clears the
 * image recolor/correction elements and the shape-level fill, line, effects,
 * and 3-D style children, while preserving the poster/media relationship,
 * crop, transform, and unknown extension XML.
 */
export const resetShapeVideoFormatting = (shape: SlideShapeData): void => {
  if (getShapeMedia(shape)?.kind !== 'video') {
    throw new Error('resetShapeVideoFormatting requires a video shape');
  }

  // Validate the media before touching the live tree. The existing image
  // reset owns the exact set of PowerPoint color-correction children.
  setShapePreset(shape, 'rect');
  resetShapeImageColorEffects(shape);

  const spPr = requireSpPr(shape);
  spPr.children = spPr.children.filter(
    (child) =>
      child.kind !== 'element' ||
      child.name.namespaceURI !== NS.dml ||
      !VIDEO_FORMAT_CHILDREN.has(child.name.localName),
  );
  commitAndRefresh(shape);
};
