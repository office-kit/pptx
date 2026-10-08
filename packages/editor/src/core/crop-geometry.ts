/**
 * Pure geometry used by the editor's picture/video Crop pane.
 *
 * OOXML stores crop sides as fractions of the source image while the shape
 * transform stores the visible frame. The editor exposes the equivalent
 * picture rectangle (size and center-relative offset) and crop rectangle.
 * Values are EMU; no XML or document mutation belongs in this module.
 */

export interface CropFrame {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface CropSides {
  readonly left: number;
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
}

export interface PictureCropGeometry {
  readonly pictureWidth: number;
  readonly pictureHeight: number;
  /** Picture center offset from the frame center. */
  readonly offsetX: number;
  readonly offsetY: number;
  /** Crop frame dimensions and absolute top-left position. */
  readonly cropWidth: number;
  readonly cropHeight: number;
  readonly cropLeft: number;
  readonly cropTop: number;
}

export interface PictureGeometryPatch {
  readonly pictureWidth?: number;
  readonly pictureHeight?: number;
  readonly offsetX?: number;
  readonly offsetY?: number;
}

export interface FrameGeometryPatch {
  readonly x?: number;
  readonly y?: number;
  readonly w?: number;
  readonly h?: number;
}

const finite = (value: number, name: string): number => {
  if (!Number.isFinite(value)) throw new RangeError(`${name} must be finite`);
  return value;
};

const positive = (value: number, name: string): number => {
  finite(value, name);
  if (value <= 0) throw new RangeError(`${name} must be positive`);
  return value;
};

const normalizedCrop = (crop: Partial<CropSides> | null | undefined): CropSides => ({
  left: finite(crop?.left ?? 0, 'crop.left'),
  top: finite(crop?.top ?? 0, 'crop.top'),
  right: finite(crop?.right ?? 0, 'crop.right'),
  bottom: finite(crop?.bottom ?? 0, 'crop.bottom'),
});

const frameChecked = (frame: CropFrame): CropFrame => ({
  x: finite(frame.x, 'frame.x'),
  y: finite(frame.y, 'frame.y'),
  w: positive(frame.w, 'frame.w'),
  h: positive(frame.h, 'frame.h'),
});

const pictureRect = (geometry: PictureCropGeometry, frame: CropFrame) => ({
  left: frame.x + frame.w / 2 + geometry.offsetX - geometry.pictureWidth / 2,
  top: frame.y + frame.h / 2 + geometry.offsetY - geometry.pictureHeight / 2,
  right: frame.x + frame.w / 2 + geometry.offsetX + geometry.pictureWidth / 2,
  bottom: frame.y + frame.h / 2 + geometry.offsetY + geometry.pictureHeight / 2,
});

const fromPictureRect = (
  frame: CropFrame,
  picture: { left: number; top: number; right: number; bottom: number },
): CropSides => ({
  left: (frame.x - picture.left) / (picture.right - picture.left),
  top: (frame.y - picture.top) / (picture.bottom - picture.top),
  right: (picture.right - (frame.x + frame.w)) / (picture.right - picture.left),
  bottom: (picture.bottom - (frame.y + frame.h)) / (picture.bottom - picture.top),
});

/** Converts OOXML crop sides and frame bounds into native Crop-pane fields. */
export function getPictureCropGeometry(
  frameInput: CropFrame,
  cropInput?: Partial<CropSides> | null,
): PictureCropGeometry {
  const frame = frameChecked(frameInput);
  const crop = normalizedCrop(cropInput);
  const sourceWidth = 1 - crop.left - crop.right;
  const sourceHeight = 1 - crop.top - crop.bottom;
  if (sourceWidth <= 0 || sourceHeight <= 0) {
    throw new RangeError('crop must leave a positive visible source area');
  }
  const pictureWidth = frame.w / sourceWidth;
  const pictureHeight = frame.h / sourceHeight;
  return {
    pictureWidth,
    pictureHeight,
    offsetX: (pictureWidth * (crop.right - crop.left)) / 2,
    offsetY: (pictureHeight * (crop.bottom - crop.top)) / 2,
    cropWidth: frame.w,
    cropHeight: frame.h,
    // The reference desktop app exposes the crop frame in slide coordinates. These are
    // deliberately the frame origin, rather than source-space crop lengths.
    cropLeft: frame.x,
    cropTop: frame.y,
  };
}

/** Applies picture size/offset edits while keeping the visible frame fixed. */
export function editPictureCropGeometry(
  frameInput: CropFrame,
  cropInput: Partial<CropSides> | null | undefined,
  patch: PictureGeometryPatch,
): { frame: CropFrame; crop: CropSides } {
  const frame = frameChecked(frameInput);
  const current = getPictureCropGeometry(frame, cropInput);
  const pictureWidth = positive(patch.pictureWidth ?? current.pictureWidth, 'pictureWidth');
  const pictureHeight = positive(patch.pictureHeight ?? current.pictureHeight, 'pictureHeight');
  const offsetX = finite(patch.offsetX ?? current.offsetX, 'offsetX');
  const offsetY = finite(patch.offsetY ?? current.offsetY, 'offsetY');
  const picture = {
    left: frame.x + frame.w / 2 + offsetX - pictureWidth / 2,
    top: frame.y + frame.h / 2 + offsetY - pictureHeight / 2,
    right: frame.x + frame.w / 2 + offsetX + pictureWidth / 2,
    bottom: frame.y + frame.h / 2 + offsetY + pictureHeight / 2,
  };
  return { frame, crop: fromPictureRect(frame, picture) };
}

/** Resets crop by expanding the frame to the picture's absolute rectangle. */
export function resetCropGeometry(
  frameInput: CropFrame,
  cropInput: Partial<CropSides> | null | undefined,
): { frame: CropFrame; crop: null } {
  const frame = frameChecked(frameInput);
  const current = getPictureCropGeometry(frame, cropInput);
  const picture = pictureRect(current, frame);
  return {
    frame: {
      x: picture.left,
      y: picture.top,
      w: picture.right - picture.left,
      h: picture.bottom - picture.top,
    },
    crop: null,
  };
}

/** Applies frame edits while keeping the picture's absolute rectangle fixed. */
export function editFrameCropGeometry(
  frameInput: CropFrame,
  cropInput: Partial<CropSides> | null | undefined,
  patch: FrameGeometryPatch,
): { frame: CropFrame; crop: CropSides } {
  const frame = frameChecked(frameInput);
  const current = getPictureCropGeometry(frame, cropInput);
  const next = frameChecked({ ...frame, ...patch });
  return { frame: next, crop: fromPictureRect(next, pictureRect(current, frame)) };
}
