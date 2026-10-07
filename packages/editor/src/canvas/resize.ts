import type { Rect } from './snapping.ts';

export type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';

const handleVectors: Record<ResizeHandle, readonly [number, number]> = {
  nw: [-1, -1],
  n: [0, -1],
  ne: [1, -1],
  e: [1, 0],
  se: [1, 1],
  s: [0, 1],
  sw: [-1, 1],
  w: [-1, 0],
};
const vectorHandles: Record<string, ResizeHandle> = {
  '-1,-1': 'nw',
  '0,-1': 'n',
  '1,-1': 'ne',
  '1,0': 'e',
  '1,1': 'se',
  '0,1': 's',
  '-1,1': 'sw',
  '-1,0': 'w',
};

/**
 * Map a source local handle into a target whose rotation differs by quarters.
 * PowerPoint snaps the relative rotation to the nearest quarter turn while
 * resizing a multi-selection, rather than continuously changing the handle.
 */
function relativeResizeHandle(handle: ResizeHandle, quarterTurns: number): ResizeHandle {
  const [x, y] = handleVectors[handle]!;
  const turns = ((quarterTurns % 4) + 4) % 4;
  if (turns === 0) return handle;
  if (turns === 1) return vectorHandles[`${y},${-x}`]!;
  if (turns === 2) return vectorHandles[`${-x},${-y}`]!;
  return vectorHandles[`${-y},${x}`]!;
}

/** Return the nearest quarter-turn, with +/-45 degree ties rounding symmetrically. */
function relativeQuarterTurns(sourceRotation: number, targetRotation: number): number {
  const relative = ((((targetRotation - sourceRotation + 180) % 360) + 360) % 360) - 180;
  return Math.sign(relative) * Math.round(Math.abs(relative) / 90);
}

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

/** PowerPoint resizes each selected object about its own opposite handle. */
export function resizeSelectionRects(
  rects: readonly (Rect & { rotation?: number })[],
  grabbed: Rect & { rotation?: number },
  handle: ResizeHandle,
  delta: { x: number; y: number },
  minimum: { w: number; h: number },
  keepAspect = false,
): Rect[] {
  const target = resizeRect(grabbed, handle, delta, grabbed.rotation ?? 0, minimum, keepAspect);
  const scaleX = grabbed.w > 0 ? target.w / grabbed.w : 1;
  const scaleY = grabbed.h > 0 ? target.h / grabbed.h : 1;
  const sourceRotation = grabbed.rotation ?? 0;
  return rects.map((rect) => {
    const rotation = rect.rotation ?? 0;
    const angle = (rotation * Math.PI) / 180;
    const quarterTurns = relativeQuarterTurns(sourceRotation, rotation);
    const targetHandle = relativeResizeHandle(handle, quarterTurns);
    const [sx, sy] = handleVectors[targetHandle]!;
    const swapsAxes = Math.abs(quarterTurns % 2) === 1;
    const targetScaleX = swapsAxes ? scaleY : scaleX;
    const targetScaleY = swapsAxes ? scaleX : scaleY;
    const dx = sx * rect.w * (targetScaleX - 1);
    const dy = sy * rect.h * (targetScaleY - 1);
    return resizeRect(
      rect,
      targetHandle,
      {
        x: dx * Math.cos(angle) - dy * Math.sin(angle),
        y: dx * Math.sin(angle) + dy * Math.cos(angle),
      },
      rotation,
      minimum,
      keepAspect,
    );
  });
}
