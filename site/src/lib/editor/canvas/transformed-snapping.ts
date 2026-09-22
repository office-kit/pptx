import { invert, project, type Matrix, type Point } from './group-space.ts';
import { snapMove, type Rect } from './snapping.ts';
import { selectionBounds, type RotatedRect } from './rotation.ts';

/** Visible envelope after the object's own rotation and all ancestor transforms. */
export function projectedBounds(rect: RotatedRect, matrix: Matrix): Rect {
  const angle = (rect.rotation * Math.PI) / 180;
  const cos = Math.cos(angle),
    sin = Math.sin(angle);
  const center = project(matrix, { x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 });
  const [a, b, c, d] = matrix;
  const width = Math.abs((a * cos + c * sin) * rect.w) + Math.abs((-a * sin + c * cos) * rect.h);
  const height = Math.abs((b * cos + d * sin) * rect.w) + Math.abs((-b * sin + d * cos) * rect.h);
  return { x: center.x - width / 2, y: center.y - height / 2, w: width, h: height };
}

/** Snap in slide coordinates, then convert only the translation back to the group. */
export function snapTransformedMove(
  moving: readonly RotatedRect[],
  others: readonly RotatedRect[],
  matrix: Matrix,
  delta: Point,
  slide: { w: number; h: number },
  threshold: number,
) {
  const inverse = invert(matrix);
  if (!inverse) return { delta, guides: [] };
  const envelope = selectionBounds(
    moving.map((rect) => ({
      ...projectedBounds({ ...rect, x: rect.x + delta.x, y: rect.y + delta.y }, matrix),
      rotation: 0,
    })),
  );
  const snapped = snapMove(
    envelope,
    others.map((rect) => projectedBounds(rect, matrix)),
    slide,
    threshold,
  );
  const origin = project(inverse, { x: 0, y: 0 });
  const correction = project(inverse, { x: snapped.x - envelope.x, y: snapped.y - envelope.y });
  return {
    delta: { x: delta.x + correction.x - origin.x, y: delta.y + correction.y - origin.y },
    guides: snapped.guides,
  };
}
