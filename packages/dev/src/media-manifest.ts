import {
  findShapesWithMedia,
  getShapeBounds,
  getShapeId,
  getShapeMedia,
  getShapeMediaPlayback,
  type MediaPlayback,
  type ShapeBounds,
  type ShapeMedia,
} from '@office-kit/pptx';
import { getSlides } from '@office-kit/pptx';

/** Media information exposed to the dev preview for a single slide shape. */
export interface PreviewMedia {
  readonly slideIndex: number;
  readonly shapeId: number;
  readonly kind: ShapeMedia['kind'];
  /** A data URL for embedded media, or the external URL for online media. */
  readonly src: string;
  /** Present for embedded media and suitable for an HTML media element. */
  readonly contentType?: string;
  /** Original shape geometry in EMUs. */
  readonly bounds: ShapeBounds;
  readonly playback: MediaPlayback | null;
}

const dataUrl = (contentType: string, bytes: Uint8Array): string =>
  `data:${contentType};base64,${Buffer.from(bytes).toString('base64')}`;

/**
 * Collects media shapes without changing the rendered SVG. Embedded clips are
 * made self-contained data URLs so the dev server does not need to expose raw
 * OPC package parts; the presentation page can replace the matching preview
 * shape with an HTML media element using the stable shape id.
 */
export const getPreviewMedia = (presentation: Parameters<typeof getSlides>[0]): PreviewMedia[] =>
  getSlides(presentation).flatMap((slide, slideIndex) =>
    findShapesWithMedia(slide).flatMap((shape) => {
      const media = getShapeMedia(shape);
      const bounds = getShapeBounds(shape);
      if (media === null || bounds === null) return [];
      const embedded = media.kind !== 'online';
      return [
        {
          slideIndex,
          shapeId: getShapeId(shape),
          kind: media.kind,
          src: embedded ? dataUrl(media.contentType, media.bytes) : media.url,
          ...(embedded ? { contentType: media.contentType } : {}),
          bounds,
          playback: embedded ? getShapeMediaPlayback(shape) : null,
        },
      ];
    }),
  );
