import {
  getShapeImageFillBytes,
  getShapeImageFormat,
  getShapeImageFillLayout,
  getShapeImageOpacity,
  getShapeImageCrop,
  setShapeImageFill,
  setShapeImageFillLayout,
  setShapeImageOpacity,
  setShapeImageCrop,
  type ImageFormat,
  type ImageFillLayout,
  type ImageCrop,
  type SlideShapeData,
} from '@office-kit/pptx';

export interface RememberedImageFill {
  bytes: Uint8Array;
  format: ImageFormat;
  layout: ImageFillLayout | null;
  opacity: number | null;
  crop: ImageCrop | null;
}

export function readRememberedImageFill(shape: SlideShapeData): RememberedImageFill | undefined {
  const bytes = getShapeImageFillBytes(shape);
  const format = getShapeImageFormat(shape);
  if (!bytes || !format) return undefined;
  const crop = getShapeImageCrop(shape);
  if (crop && Object.values(crop).some((value) => !Number.isFinite(value))) return undefined;
  return {
    bytes: bytes.slice(),
    format,
    layout: getShapeImageFillLayout(shape),
    opacity: getShapeImageOpacity(shape),
    crop,
  };
}

export function restoreRememberedImageFill(shape: SlideShapeData, fill: RememberedImageFill) {
  setShapeImageFill(shape, fill.bytes, { format: fill.format });
  if (fill.layout) setShapeImageFillLayout(shape, fill.layout);
  setShapeImageOpacity(shape, fill.opacity);
  setShapeImageCrop(shape, fill.crop);
}

export interface RememberedImageLayouts {
  tile?: Extract<ImageFillLayout, { mode: 'tile' }>;
  stretch?: Extract<ImageFillLayout, { mode: 'stretch' }>;
}

export function switchRememberedImageLayout(
  current: ImageFillLayout,
  mode: ImageFillLayout['mode'],
  remembered: RememberedImageLayouts,
): ImageFillLayout {
  if (current.mode === 'tile') remembered.tile = { ...current };
  else remembered.stretch = { ...current };
  return { ...(remembered[mode] ?? { mode }), rotateWithShape: current.rotateWithShape };
}
