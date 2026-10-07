import {
  getGroupChildren,
  getGroupTransform,
  getShapeFlip,
  getShapeId,
  getShapeRotation,
  getSlideShapes,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';

/** Affine coordinates in EMU, using the same order as SVG/CSS matrix(). */
export type Matrix = readonly [number, number, number, number, number, number];
export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
export interface Point {
  readonly x: number;
  readonly y: number;
}

export function compose(parent: Matrix, child: Matrix): Matrix {
  const [a, b, c, d, e, f] = parent;
  const [g, h, i, j, k, l] = child;
  return [
    a * g + c * h,
    b * g + d * h,
    a * i + c * j,
    b * i + d * j,
    a * k + c * l + e,
    b * k + d * l + f,
  ];
}
export function project(matrix: Matrix, point: Point): Point {
  const [a, b, c, d, e, f] = matrix;
  return { x: a * point.x + c * point.y + e, y: b * point.x + d * point.y + f };
}
export function invert(matrix: Matrix): Matrix | null {
  const [a, b, c, d, e, f] = matrix;
  const determinant = a * d - b * c;
  if (!Number.isFinite(determinant) || determinant === 0) return null;
  const result: Matrix = [
    d / determinant,
    -b / determinant,
    -c / determinant,
    a / determinant,
    (c * f - d * e) / determinant,
    (b * e - a * f) / determinant,
  ];
  return result.every(Number.isFinite) ? result : null;
}

/** Internal child coordinates → the group's parent coordinate system. */
export function groupMatrix(group: SlideShapeData): Matrix {
  const transform = getGroupTransform(group);
  if (!transform) return IDENTITY;
  const { outer, inner } = transform;
  const flip = getShapeFlip(group);
  const sx = (outer.w / (inner.w || 1)) * (flip?.horizontal ? -1 : 1);
  const sy = (outer.h / (inner.h || 1)) * (flip?.vertical ? -1 : 1);
  const angle = (getShapeRotation(group) * Math.PI) / 180;
  const cos = Math.cos(angle),
    sin = Math.sin(angle);
  const a = cos * sx,
    b = sin * sx,
    c = -sin * sy,
    d = cos * sy;
  const ix = inner.x + (inner.w || 1) / 2,
    iy = inner.y + (inner.h || 1) / 2;
  return [
    a,
    b,
    c,
    d,
    outer.x + outer.w / 2 - a * ix - c * iy,
    outer.y + outer.h / 2 - b * ix - d * iy,
  ];
}

export interface ShapeScope {
  readonly parent: SlideShapeData | null;
  readonly shapes: readonly SlideShapeData[];
  /** Scope-local coordinates → slide coordinates, including all ancestors. */
  readonly matrix: Matrix;
  readonly textScale: { readonly x: number; readonly y: number };
}

/** Select siblings in their original stacking order, with an iterative ancestor walk. */
export function shapeScope(slide: SlideData, selectedId: number | null): ShapeScope {
  const shapes = getSlideShapes(slide);
  const parents = new Map<number, SlideShapeData>();
  const children = new Map<number, readonly SlideShapeData[]>();
  for (const shape of shapes) {
    const members = getGroupChildren(shape);
    if (!members.length) continue;
    children.set(getShapeId(shape), members);
    for (const child of members) parents.set(getShapeId(child), shape);
  }
  const parent = selectedId === null ? null : (parents.get(selectedId) ?? null);
  let matrix = IDENTITY;
  let scaleX = 1,
    scaleY = 1;
  let ancestor = parent;
  while (ancestor) {
    const transform = getGroupTransform(ancestor);
    if (transform) {
      scaleX *= transform.outer.w / (transform.inner.w || 1);
      scaleY *= transform.outer.h / (transform.inner.h || 1);
    }
    matrix = compose(groupMatrix(ancestor), matrix);
    ancestor = parents.get(getShapeId(ancestor)) ?? null;
  }
  return {
    parent,
    textScale: { x: scaleX, y: scaleY },
    shapes: parent
      ? children.get(getShapeId(parent))!
      : shapes.filter((shape) => !parents.has(getShapeId(shape))),
    matrix,
  };
}
