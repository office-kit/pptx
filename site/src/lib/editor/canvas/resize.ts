import type { Rect } from './snapping.ts';

export type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

/** Resize in the shape's local axes, keeping the opposite handle fixed on the slide. */
export function resizeRect(
  rect: Rect,
  handle: ResizeHandle,
  delta: { x: number; y: number },
  rotation: number,
  minimum: { w: number; h: number },
  keepAspect: boolean,
): Rect {
  const angle = (rotation * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const dx = delta.x * cos + delta.y * sin;
  const dy = -delta.x * sin + delta.y * cos;
  const sx = handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0;
  const sy = handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0;
  let w = rect.w + sx * dx;
  let h = rect.h + sy * dy;
  if (keepAspect && rect.w > 0 && rect.h > 0) {
    const scaleX = w / rect.w;
    const scaleY = h / rect.h;
    const scale = Math.max(
      minimum.w / rect.w,
      minimum.h / rect.h,
      !sx ? scaleY : !sy ? scaleX : Math.abs(scaleX - 1) >= Math.abs(scaleY - 1) ? scaleX : scaleY,
    );
    w = rect.w * scale;
    h = rect.h * scale;
  } else {
    if (sx) w = Math.max(minimum.w, w);
    if (sy) h = Math.max(minimum.h, h);
  }
  // Transform the local centre shift back to the slide; unrotated x/y describe
  // the new box around that centre, not the visual top-left of a rotated shape.
  const cx = (sx * (w - rect.w)) / 2;
  const cy = (sy * (h - rect.h)) / 2;
  return {
    x: rect.x + (rect.w - w) / 2 + cx * cos - cy * sin,
    y: rect.y + (rect.h - h) / 2 + cx * sin + cy * cos,
    w,
    h,
  };
}

/** Uniformly scale a selection without introducing shear into rotated objects. */
export function resizeSelectionRects(
  rects: readonly Rect[],
  frame: Rect,
  handle: ResizeHandle,
  delta: { x: number; y: number },
  minimum: { w: number; h: number },
): Rect[] {
  const target = resizeRect(frame, handle, delta, 0, minimum, true);
  const scale = frame.w > 0 ? target.w / frame.w : frame.h > 0 ? target.h / frame.h : 1;
  const anchor = {
    x: frame.x + (handle.includes('w') ? frame.w : handle.includes('e') ? 0 : frame.w / 2),
    y: frame.y + (handle.includes('n') ? frame.h : handle.includes('s') ? 0 : frame.h / 2),
  };
  return rects.map((rect) => ({
    x: anchor.x + (rect.x - anchor.x) * scale,
    y: anchor.y + (rect.y - anchor.y) * scale,
    w: rect.w * scale,
    h: rect.h * scale,
  }));
}
