import {
  setShapeImageFill,
  getShapeImageFillLayout,
  getShapeImageOpacity,
  setShapeImageFillLayout,
  setShapeImageOpacity,
  getShapeFillEffective,
  getShapeFillOpacity,
  getShapeGradientFillEffective,
  getShapePatternFill,
  getShapeFillColorResolved,
  asColor,
  emu,
  type ImageFillLayout,
  type PresentationData,
  type SlideData,
  type SlideShapeData,
  type Color,
  type GradientFillOptions,
  type PatternFillOptions,
} from '@office-kit/pptx';
import {
  readRememberedImageFill,
  switchRememberedImageLayout,
  type RememberedImageFill,
  type RememberedImageLayouts,
} from './remembered-image-fill.ts';

export interface RememberedFill {
  solid?: { color: Color; opacity?: number };
  gradient?: GradientFillOptions;
  pattern?: PatternFillOptions;
  image?: RememberedImageFill;
  imageLayouts?: RememberedImageLayouts;
  backgroundImage?: SlideData;
}

export function rememberShapeFill(
  pres: PresentationData,
  target: SlideShapeData,
  remembered: RememberedFill,
): void {
  const current = getShapeFillEffective(pres, target);
  if (current.kind === 'image') remembered.image = readRememberedImageFill(pres, target);
  if (current.kind === 'pattern')
    remembered.pattern = getShapePatternFill(pres, target, { preserveTheme: true }) ?? undefined;
  if (current.kind === 'solid')
    remembered.solid = {
      color: asColor(getShapeFillColorResolved(pres, target) ?? current.color) ?? 'accent1',
      opacity: getShapeFillOpacity(target) ?? undefined,
    };
  if (current.kind === 'gradient') {
    const gradient = getShapeGradientFillEffective(pres, target);
    if (gradient)
      remembered.gradient = {
        ...gradient,
        stops: gradient.stops.map((stop) => {
          const color = asColor(stop.color);
          return {
            ...stop,
            color: color ?? asColor(stop.resolvedColor ?? '') ?? 'accent1',
            brightness: color ? stop.brightness : 0,
          };
        }),
      };
  }
}

export function insertRememberedPictureFill(
  pres: PresentationData,
  target: SlideShapeData,
  bytes: Uint8Array,
  remembered: RememberedFill,
): void {
  rememberShapeFill(pres, target, remembered);
  const current = getShapeImageFillLayout(target) ?? remembered.image?.layout;
  const opacity = getShapeImageOpacity(target) ?? remembered.image?.opacity ?? null;
  const layout = current
    ? switchRememberedImageLayout(current, 'stretch', (remembered.imageLayouts ??= {}))
    : { mode: 'stretch' as const };
  setShapeImageFill(target, bytes);
  setShapeImageFillLayout(target, layout);
  setShapeImageOpacity(target, opacity);
}

/** The reference desktop app's tiling for a gallery texture: no offset, 100%, top left, no mirror, rotating with the shape. */
export const TEXTURE_TILE_LAYOUT = {
  mode: 'tile',
  offsetX: emu(0),
  offsetY: emu(0),
  scaleX: 1,
  scaleY: 1,
  alignment: 'tl',
  flip: 'none',
  rotateWithShape: true,
} as const satisfies ImageFillLayout;

/** Fills the shape with a gallery texture, keeping its picture transparency like a picture insert does. */
export function insertRememberedTextureFill(
  pres: PresentationData,
  target: SlideShapeData,
  bytes: Uint8Array,
  remembered: RememberedFill,
): void {
  rememberShapeFill(pres, target, remembered);
  const current = getShapeImageFillLayout(target);
  // Record the outgoing placement so unticking Tile restores its stretch offsets.
  if (current) switchRememberedImageLayout(current, 'tile', (remembered.imageLayouts ??= {}));
  const opacity = getShapeImageOpacity(target) ?? remembered.image?.opacity ?? null;
  setShapeImageFill(target, bytes);
  setShapeImageFillLayout(target, TEXTURE_TILE_LAYOUT);
  setShapeImageOpacity(target, opacity);
}
