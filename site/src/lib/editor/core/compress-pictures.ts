// Compress Pictures: downsample embedded raster pictures to a target
// resolution and, optionally, drop their cropped-away pixels.
import {
  getGroupChildren,
  getGroupTransform,
  getShapeBounds,
  getShapeImageBytes,
  getShapeImageCrop,
  getShapeImageFormat,
  getShapeId,
  getShapeKind,
  getShapeMedia,
  getSlideShapes,
  getSlides,
  getShapeXmlString,
  type ImageCompressionState,
  type ImageCrop,
  type PresentationData,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';

/**
 * Mac PowerPoint's Compress Pictures ▸ Picture Quality choices, in menu order,
 * with the `a:blip/@cstate` each one writes. PowerPoint 16.113 labels a
 * picture only for Print, On-screen and Email; High Fidelity, HD and Use
 * Original Quality leave its blip as it was.
 */
export const PICTURE_QUALITIES = [
  { id: 'highFidelity', label: 'High Fidelity (Maximum ppi)', ppi: null, cstate: null },
  { id: 'hd', label: 'HD (330 ppi)', ppi: 330, cstate: null },
  { id: 'print', label: 'Print (220 ppi)', ppi: 220, cstate: 'print' },
  { id: 'screen', label: 'On-screen (150 ppi)', ppi: 150, cstate: 'screen' },
  { id: 'email', label: 'Email (96 ppi)', ppi: 96, cstate: 'email' },
  { id: 'original', label: 'Use Original Quality', ppi: null, cstate: null },
] as const satisfies readonly {
  id: string;
  label: string;
  ppi: number | null;
  cstate: ImageCompressionState | null;
}[];
export type PictureQuality = (typeof PICTURE_QUALITIES)[number]['id'];

const EMU_PER_INCH = 914400;

/** A picture on a slide, with the scale its groups apply to its frame. */
export interface PictureTarget {
  readonly slide: SlideData;
  /** The picture's id followed by the ids of the groups holding it. */
  readonly ids: readonly number[];
  readonly shape: SlideShapeData;
  readonly scaleX: number;
  readonly scaleY: number;
}

/** Every picture in the presentation (videos and audio excluded), groups included. */
export function collectPictures(pres: PresentationData): PictureTarget[] {
  const out: PictureTarget[] = [];
  for (const slide of getSlides(pres)) {
    // getSlideShapes lists group members too; walk from the top-level shapes
    // to learn each member's group scale, and mutate through those handles.
    const shapes = getSlideShapes(slide);
    const byId = new Map(shapes.map((shape) => [getShapeId(shape), shape]));
    const members = new Set(shapes.flatMap((shape) => getGroupChildren(shape).map(getShapeId)));
    const visit = (
      shape: SlideShapeData,
      ids: readonly number[],
      scaleX: number,
      scaleY: number,
    ) => {
      const id = getShapeId(shape);
      const handle = byId.get(id) ?? shape;
      const kind = getShapeKind(handle);
      if (kind === 'picture' && !getShapeMedia(handle))
        out.push({ slide, ids: [id, ...ids], shape: handle, scaleX, scaleY });
      if (kind !== 'group') return;
      const transform = getGroupTransform(handle);
      const sx = transform && transform.inner.w > 0 ? transform.outer.w / transform.inner.w : 1;
      const sy = transform && transform.inner.h > 0 ? transform.outer.h / transform.inner.h : 1;
      for (const child of getGroupChildren(handle))
        visit(child, [id, ...ids], scaleX * sx, scaleY * sy);
    };
    for (const shape of shapes) if (!members.has(getShapeId(shape))) visit(shape, [], 1, 1);
  }
  return out;
}

export interface CompressionInput {
  readonly pixelWidth: number;
  readonly pixelHeight: number;
  /** The frame's size on the slide, in EMU. */
  readonly frameWidth: number;
  readonly frameHeight: number;
  readonly crop: ImageCrop | null;
  /** Target resolution, or null to keep every pixel. */
  readonly ppi: number | null;
  readonly deleteCropped: boolean;
}

export interface CompressionPlan {
  /** Source pixels to keep. */
  readonly source: {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
  };
  /** Output size in pixels. */
  readonly width: number;
  readonly height: number;
  /** The crop to write once the new pixels replace the old ones. */
  readonly crop: ImageCrop | null;
}

const side = (crop: ImageCrop | null, key: keyof ImageCrop) => crop?.[key] ?? 0;

/**
 * What compressing one picture does, or null when it would change nothing.
 * Never upsamples. A negative crop side (the frame extends past the picture)
 * survives cropped-area deletion, rescaled to the smaller picture.
 */
export function planCompression(input: CompressionInput): CompressionPlan | null {
  const { pixelWidth: w, pixelHeight: h } = input;
  const [l, t, r, b] = (['left', 'top', 'right', 'bottom'] as const).map((key) =>
    side(input.crop, key),
  ) as [number, number, number, number];
  const visibleX = 1 - l - r;
  const visibleY = 1 - t - b;
  if (
    w <= 0 ||
    h <= 0 ||
    visibleX <= 0 ||
    visibleY <= 0 ||
    input.frameWidth <= 0 ||
    input.frameHeight <= 0
  )
    return null;
  const cut = input.deleteCropped && (l > 0 || t > 0 || r > 0 || b > 0);
  const [cl, ct, cr, cb] = cut
    ? ([l, t, r, b].map((value) => Math.max(value, 0)) as [number, number, number, number])
    : [0, 0, 0, 0];
  const keptX = 1 - cl - cr;
  const keptY = 1 - ct - cb;
  const x0 = Math.round(cl * w);
  const y0 = Math.round(ct * h);
  const source = {
    x: x0,
    y: y0,
    width: Math.max(1, Math.round((cl + keptX) * w) - x0),
    height: Math.max(1, Math.round((ct + keptY) * h) - y0),
  };
  // Pixels per inch of the visible part of the picture, per axis. Keep both
  // axes at or above the target so neither direction loses detail it needs.
  let scale = 1;
  if (input.ppi !== null) {
    const ppiX = (w * visibleX) / (input.frameWidth / EMU_PER_INCH);
    const ppiY = (h * visibleY) / (input.frameHeight / EMU_PER_INCH);
    scale = Math.min(1, Math.max(input.ppi / ppiX, input.ppi / ppiY));
  }
  const width = Math.max(1, Math.round(source.width * scale));
  const height = Math.max(1, Math.round(source.height * scale));
  if (!cut && width === w && height === h) return null;
  const rescale = (value: number, kept: number) => (value < 0 ? value / kept : 0);
  const crop: ImageCrop | null = cut
    ? {
        left: rescale(l, keptX),
        top: rescale(t, keptY),
        right: rescale(r, keptX),
        bottom: rescale(b, keptY),
      }
    : input.crop;
  const empty = crop && !crop.left && !crop.top && !crop.right && !crop.bottom;
  return { source, width, height, crop: empty ? null : crop };
}

// Browsers' default JPEG quality is lower than PowerPoint's re-encodes.
const JPEG_QUALITY = 0.9;
// [MS-ODRAWXML] `imgProps`: the blip extension holding PowerPoint's picture-editing
// data (Artistic Effects and corrections) and its JPEG XR original.
const IMAGE_EDITING_DATA_URI = /uri="\{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B\}"/i;

export interface CompressedPicture {
  readonly target: PictureTarget;
  readonly bytes: Uint8Array;
  readonly crop: ImageCrop | null;
}

/**
 * Whether resampled bytes replace the picture. A crop deletion always does.
 * A resolution change alone only does when the result is smaller: PowerPoint
 * 16.113 resampled a 1200 px PNG shown at 240 ppi for Email (to 480 px, the
 * frame's 5 in × 96 ppi, and only 11% fewer bytes: 55978 → 50079) but left it
 * untouched for Print (220 ppi) and On-screen (150 ppi), although both are
 * below 240 ppi. Its 480 px re-encode kept 89% of the bytes for 16% of the
 * pixels, so a 1100 or 750 px one would almost certainly have outgrown the
 * original. PowerPoint does not document its rule; keeping whichever is
 * smaller is the simplest one that matches every captured case.
 */
export const keepsResampledBytes = (
  plan: CompressionPlan,
  pixelWidth: number,
  pixelHeight: number,
  originalBytes: number,
  resampledBytes: number,
): boolean =>
  plan.source.width !== pixelWidth ||
  plan.source.height !== pixelHeight ||
  resampledBytes < originalBytes;

/**
 * Resamples each PNG or JPEG picture whose plan changes it. Other formats
 * (vector, GIF, TIFF …) are left alone. A picture carrying PowerPoint's
 * picture-editing data keeps its crop: its JPEG XR original cannot be cropped
 * to match here, and PowerPoint re-renders the picture from that original.
 */
export async function compressPictures(
  targets: readonly PictureTarget[],
  ppi: number | null,
  deleteCropped: boolean,
): Promise<CompressedPicture[]> {
  const out: CompressedPicture[] = [];
  // One picture at a time keeps a single decoded bitmap in memory.
  for (const target of targets) {
    const format = getShapeImageFormat(target.shape);
    const bytes = getShapeImageBytes(target.shape);
    const bounds = getShapeBounds(target.shape);
    if ((format !== 'png' && format !== 'jpeg') || !bytes || !bounds) continue;
    const type = format === 'png' ? 'image/png' : 'image/jpeg';
    const bitmap = await createImageBitmap(new Blob([bytes.slice()], { type }));
    try {
      const plan = planCompression({
        pixelWidth: bitmap.width,
        pixelHeight: bitmap.height,
        frameWidth: bounds.w * target.scaleX,
        frameHeight: bounds.h * target.scaleY,
        crop: getShapeImageCrop(target.shape),
        ppi,
        deleteCropped:
          deleteCropped && !IMAGE_EDITING_DATA_URI.test(getShapeXmlString(target.shape)),
      });
      if (!plan) continue;
      const canvas = new OffscreenCanvas(plan.width, plan.height);
      const context = canvas.getContext('2d')!;
      context.imageSmoothingQuality = 'high';
      const { x, y, width, height } = plan.source;
      context.drawImage(bitmap, x, y, width, height, 0, 0, plan.width, plan.height);
      const blob = await canvas.convertToBlob(
        type === 'image/jpeg' ? { type, quality: JPEG_QUALITY } : { type },
      );
      if (keepsResampledBytes(plan, bitmap.width, bitmap.height, bytes.byteLength, blob.size))
        out.push({ target, bytes: new Uint8Array(await blob.arrayBuffer()), crop: plan.crop });
    } finally {
      bitmap.close();
    }
  }
  return out;
}
