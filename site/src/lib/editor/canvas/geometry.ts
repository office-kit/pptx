// Canvas ↔ slide coordinate helpers. The slide is authored in EMU; the canvas
// paints it at some pixel size. These convert between the two and express a
// shape's box as percentages of the slide so overlays stay aligned at any
// zoom without re-measuring.

import {
  getShapeBoundsResolved,
  getShapeId,
  getShapeRotation,
  getSlideSize,
} from '@office-kit/pptx';
import type { PresentationData, SlideData, SlideShapeData } from '@office-kit/pptx';

export const DEFAULT_SLIDE = { width: 12192000, height: 6858000 };

export interface Box {
  readonly id: number;
  /** Percentages of the slide (0–100). */
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  readonly rotation: number;
  readonly shape: SlideShapeData;
}

export interface SlideMetrics {
  readonly widthEmu: number;
  readonly heightEmu: number;
}

export function slideMetrics(pres: PresentationData): SlideMetrics {
  const size = getSlideSize(pres);
  return {
    widthEmu: (size?.width as unknown as number) ?? DEFAULT_SLIDE.width,
    heightEmu: (size?.height as unknown as number) ?? DEFAULT_SLIDE.height,
  };
}

export function shapeBoxes(
  pres: PresentationData,
  slide: SlideData,
  shapes: readonly SlideShapeData[],
): Box[] {
  const m = slideMetrics(pres);
  const out: Box[] = [];
  for (const shape of shapes) {
    const b = getShapeBoundsResolved(pres, shape);
    if (!b) continue;
    out.push({
      id: getShapeId(shape),
      left: ((b.x as unknown as number) / m.widthEmu) * 100,
      top: ((b.y as unknown as number) / m.heightEmu) * 100,
      width: ((b.w as unknown as number) / m.widthEmu) * 100,
      height: ((b.h as unknown as number) / m.heightEmu) * 100,
      rotation: safeRotation(shape),
      shape,
    });
  }
  return out;
}

function safeRotation(shape: SlideShapeData): number {
  try {
    return getShapeRotation(shape);
  } catch {
    return 0;
  }
}

/** EMU-per-pixel for a stage rendered at `rectW × rectH` pixels. */
export function emuPerPx(m: SlideMetrics, rectW: number, rectH: number) {
  return {
    x: rectW ? m.widthEmu / rectW : 1,
    y: rectH ? m.heightEmu / rectH : 1,
  };
}
