import type { Rect } from './snapping.ts';

export interface RotatedRect extends Rect {
  rotation: number;
}

/** Axis-aligned envelope of the visible, rotated rectangles. */
export function selectionBounds(rects: readonly RotatedRect[]): Rect {
  let left = Infinity,
    top = Infinity,
    right = -Infinity,
    bottom = -Infinity;
  for (const rect of rects) {
    const angle = (rect.rotation * Math.PI) / 180;
    const w = Math.abs(rect.w * Math.cos(angle)) + Math.abs(rect.h * Math.sin(angle));
    const h = Math.abs(rect.w * Math.sin(angle)) + Math.abs(rect.h * Math.cos(angle));
    const cx = rect.x + rect.w / 2,
      cy = rect.y + rect.h / 2;
    left = Math.min(left, cx - w / 2);
    top = Math.min(top, cy - h / 2);
    right = Math.max(right, cx + w / 2);
    bottom = Math.max(bottom, cy + h / 2);
  }
  return { x: left, y: top, w: right - left, h: bottom - top };
}

/** Rotate centres and orientations together, preserving sizes and spacing. */
export function rotateRect(
  rect: RotatedRect,
  centre: { x: number; y: number },
  degrees: number,
): RotatedRect {
  const angle = (degrees * Math.PI) / 180;
  const x = rect.x + rect.w / 2 - centre.x,
    y = rect.y + rect.h / 2 - centre.y;
  return {
    x: centre.x + x * Math.cos(angle) - y * Math.sin(angle) - rect.w / 2,
    y: centre.y + x * Math.sin(angle) + y * Math.cos(angle) - rect.h / 2,
    w: rect.w,
    h: rect.h,
    rotation: (((rect.rotation + degrees) % 360) + 360) % 360,
  };
}
