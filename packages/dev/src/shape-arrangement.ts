import type { EditorShape } from './editor.ts';
import { shapePoint } from './connector-geometry.ts';

type Box = NonNullable<EditorShape['bounds']>;

/** Axis-aligned extent in slide coordinates, including all ancestor transforms. */
export function slideExtent(shape: EditorShape): Box | null {
  const box = shape.bounds,
    m = shape.parentTransform;
  if (!box || m === null) return null;
  if (m && Math.abs(m[0] * m[3] - m[1] * m[2]) < 1e-12) return null;
  const points = [
    [0, 0],
    [box.w, 0],
    [0, box.h],
    [box.w, box.h],
  ].map(([x, y]) => {
    const p = shapePoint(shape, { x: x!, y: y! });
    return m ? { x: m[0] * p.x + m[2] * p.y + m[4], y: m[1] * p.x + m[3] * p.y + m[5] } : p;
  });
  const x = Math.min(...points.map((p) => p.x)),
    y = Math.min(...points.map((p) => p.y));
  return {
    x,
    y,
    w: Math.max(...points.map((p) => p.x)) - x,
    h: Math.max(...points.map((p) => p.y)) - y,
  };
}
function translate(shape: EditorShape, dx: number, dy: number) {
  const m = shape.parentTransform;
  const det = m ? m[0] * m[3] - m[1] * m[2] : 1;
  return {
    id: shape.id,
    bounds: {
      ...shape.bounds!,
      x: Math.round(shape.bounds!.x + (m ? (m[3] * dx - m[2] * dy) / det : dx)),
      y: Math.round(shape.bounds!.y + (m ? (m[0] * dy - m[1] * dx) / det : dy)),
    },
  };
}
const measured = (shapes: EditorShape[]) =>
  shapes.flatMap((shape) => {
    const box = slideExtent(shape);
    return box ? [{ shape, box }] : [];
  });
export function alignedPositions(
  shapes: EditorShape[],
  direction: number,
  size: { width: number; height: number },
  toSlide = shapes.length === 1,
) {
  const items = measured(shapes);
  if (!items.length) return [];
  const left = toSlide ? 0 : Math.min(...items.map((i) => i.box.x));
  const top = toSlide ? 0 : Math.min(...items.map((i) => i.box.y));
  const right = toSlide ? size.width : Math.max(...items.map((i) => i.box.x + i.box.w));
  const bottom = toSlide ? size.height : Math.max(...items.map((i) => i.box.y + i.box.h));
  return items.map(({ shape, box }) => {
    const x =
      direction === 0
        ? left
        : direction === 1
          ? (left + right - box.w) / 2
          : direction === 2
            ? right - box.w
            : box.x;
    const y =
      direction === 3
        ? top
        : direction === 4
          ? (top + bottom - box.h) / 2
          : direction === 5
            ? bottom - box.h
            : box.y;
    return translate(shape, x - box.x, y - box.y);
  });
}
export function distributedPositions(
  shapes: EditorShape[],
  axis: 'x' | 'y',
  slideSize?: { width: number; height: number },
) {
  const items = measured(shapes).sort((a, b) => a.box[axis] - b.box[axis]);
  if (items.length < (slideSize ? 2 : 3)) return [];
  const extent = axis === 'x' ? 'w' : 'h';
  const first = items[0]!.box,
    last = items.at(-1)!.box;
  const total = items.reduce((sum, i) => sum + i.box[extent], 0);
  const gap = slideSize
    ? ((axis === 'x' ? slideSize.width : slideSize.height) - total) / (items.length + 1)
    : (last[axis] + last[extent] - first[axis] - total) / (items.length - 1);
  let position = slideSize ? gap : first[axis];
  return items.map(({ shape, box }) => {
    const delta = position - box[axis];
    position += box[extent] + gap;
    return translate(shape, axis === 'x' ? delta : 0, axis === 'y' ? delta : 0);
  });
}
