import { readImagePixelSize, readImageResolution } from '../../internal/opc/image-format.ts';
import { getShapeImageBytes, getShapeImageFillBytes } from './shape-image-effects.ts';
import {
  NS,
  firstChildElement,
  getAttrValue,
  qname,
  type XmlElement,
} from '../../internal/xml/index.ts';
import { SHAPE_ELEMENT, SHAPE_SNAPSHOT, type SlideShapeData } from '../_internal-symbols.ts';
import { inches, type Emu } from '../units.ts';
import { commitAndRefresh } from './_helpers.ts';

export type { ImageFillLayout, ImageTileAlignment, ImageTileFlip } from './_image-fill-layout.ts';
import {
  readImageFillLayout,
  writeImageFillLayout,
  type ImageFillLayout,
} from './_image-fill-layout.ts';

const fillElement = (shape: SlideShapeData): XmlElement | null => {
  const element = shape[SHAPE_ELEMENT];
  if (shape[SHAPE_SNAPSHOT].kind === 'picture') {
    return firstChildElement(element, qname('p', 'blipFill', NS.pml));
  }
  const spPr = firstChildElement(element, qname('p', 'spPr', NS.pml));
  return spPr ? firstChildElement(spPr, qname('a', 'blipFill', NS.dml)) : null;
};
/** Reads direct picture/image-fill placement; returns null when there is no image fill. */
export const getShapeImageFillLayout = (shape: SlideShapeData): ImageFillLayout | null => {
  const fill = fillElement(shape);
  return fill ? readImageFillLayout(fill) : null;
};

/** Changes placement without replacing media, crop, or effects. Invalid inputs leave the shape unchanged. */
export const setShapeImageFillLayout = (shape: SlideShapeData, layout: ImageFillLayout): void => {
  const fill = fillElement(shape);
  if (!fill)
    throw new Error('setShapeImageFillLayout requires a picture or a shape with an image fill');
  writeImageFillLayout(fill, layout);
  commitAndRefresh(shape);
};

/**
 * Natural image-fill size before tile scaling, in EMUs. Reads PNG/JPEG dimensions,
 * the fill's DPI override, and PNG pHYs/JPEG JFIF resolution. Missing resolution
 * uses 96 DPI. Returns null for unavailable bytes or unsupported image headers.
 */
export const getShapeImageIntrinsicSize = (
  shape: SlideShapeData,
): { width: Emu; height: Emu } | null => {
  const fill = fillElement(shape);
  if (!fill) return null;
  const bytes = getShapeImageBytes(shape) ?? getShapeImageFillBytes(shape);
  if (!bytes) return null;
  const pixels = readImagePixelSize(bytes);
  if (!pixels) return null;
  const override = Number(getAttrValue(fill, qname('', 'dpi', '')));
  const resolution = readImageResolution(bytes);
  const dpiX = override > 0 ? override : (resolution?.x ?? 96);
  const dpiY = override > 0 ? override : (resolution?.y ?? 96);
  return {
    width: inches(pixels.width / dpiX),
    height: inches(pixels.height / dpiY),
  };
};
