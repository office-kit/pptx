import {
  addSlideLine,
  addSlideShape,
  emu,
  getShapeBounds,
  getShapeCustomGeometry,
  getShapeId,
  getMaxShapeId,
  getShapeName,
  getSlideShapes,
  setShapeCustomGeometry,
  setShapeNoFill,
  setShapeStroke,
  setShapeStrokeCap,
  setShapeStrokeJoin,
  type SlideData,
  type SlideShapeData,
} from '@office-kit/pptx';

export interface Pen {
  readonly id: string;
  readonly label: string;
  readonly color: `#${string}`;
  readonly widthEmu: number;
  readonly opacity: number;
}

// Mac PowerPoint's default Draw gallery: black pen, red pen, pencil and a
// yellow highlighter.
export const PENS: readonly Pen[] = [
  { id: 'black', label: 'Pen: Black, 1 mm', color: '#000000', widthEmu: 36000, opacity: 1 },
  { id: 'red', label: 'Pen: Red, 1 mm', color: '#C00000', widthEmu: 36000, opacity: 1 },
  { id: 'pencil', label: 'Pencil: Gray, 1 mm', color: '#7F7F7F', widthEmu: 36000, opacity: 0.8 },
  {
    id: 'highlighter',
    label: 'Highlighter: Yellow, 6 mm',
    color: '#FFFF00',
    widthEmu: 216000,
    opacity: 0.5,
  },
];

export type InkTool = 'pen' | 'eraser' | 'lasso';

export class InkState {
  tool = $state<InkTool | null>(null);
  pens = $state<Pen[]>([...PENS]);
  pen = $state<Pen>(PENS[0]!);
  /** Ink to Shape: each pen stroke that reads as a line or simple outline becomes that shape. */
  toShape = $state(false);

  toggle(tool: InkTool): void {
    this.tool = this.tool === tool ? null : tool;
  }
  choosePen(pen: Pen): void {
    this.pen = pen;
    this.tool = 'pen';
  }
  addPen(color: `#${string}`): void {
    const pen = {
      id: `custom-${this.pens.length}`,
      label: `Pen: ${color.toUpperCase()}, 1 mm`,
      color,
      widthEmu: 36000,
      opacity: 1,
    };
    this.pens = [...this.pens, pen];
    this.choosePen(pen);
  }
}

export interface Point {
  readonly x: number;
  readonly y: number;
}

// PowerPoint names freehand strokes "Ink N"; the eraser only removes those, so
// it never deletes a freeform the author drew with the Shapes gallery.
const INK_NAME = /^Ink \d+$/;
export const isInk = (shape: SlideShapeData): boolean => INK_NAME.test(getShapeName(shape));

/** Adds one pen stroke (points in slide EMU) as an open freeform. */
export function addInkStroke(slide: SlideData, points: readonly Point[], pen: Pen): SlideShapeData {
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const x = Math.round(Math.min(...xs));
  const y = Math.round(Math.min(...ys));
  // A dot is a zero-size box; the path space must still be positive.
  const w = Math.max(1, Math.round(Math.max(...xs)) - x);
  const h = Math.max(1, Math.round(Math.max(...ys)) - y);
  const shape = addSlideShape(slide, {
    preset: 'rect',
    x: emu(x),
    y: emu(y),
    w: emu(w),
    h: emu(h),
    name: `Ink ${getMaxShapeId(slide) + 1}`,
  });
  const local = points.map((point) => ({ x: point.x - x, y: point.y - y }));
  setShapeCustomGeometry(shape, {
    paths: [
      {
        w,
        h,
        fill: 'none',
        commands: [
          { kind: 'moveTo', pt: local[0]! },
          // A single click still leaves a visible dot.
          ...(local.length === 1 ? [local[0]!] : local.slice(1)).map((pt) => ({
            kind: 'lnTo' as const,
            pt,
          })),
        ],
      },
    ],
  });
  setShapeNoFill(shape);
  setShapeStroke(shape, { color: pen.color, widthEmu: pen.widthEmu, opacity: pen.opacity });
  setShapeStrokeCap(shape, 'rnd');
  setShapeStrokeJoin(shape, 'round');
  return shape;
}

function distanceToSegment(p: Point, a: Point, b: Point): number {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const length = dx * dx + dy * dy;
  const t =
    length === 0 ? 0 : Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / length));
  return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy));
}

/** The ink stroke on the slide that passes within `tolerance` EMU of `point`, topmost first. */
export function inkAt(slide: SlideData, point: Point, tolerance: number): SlideShapeData | null {
  const shapes = getSlideShapes(slide).filter(isInk);
  for (const shape of shapes.reverse()) {
    const bounds = getShapeBounds(shape);
    const geometry = getShapeCustomGeometry(shape);
    if (!bounds || !geometry) continue;
    for (const path of geometry.paths) {
      const sx = path.w ? bounds.w / path.w : 1;
      const sy = path.h ? bounds.h / path.h : 1;
      const points = path.commands.flatMap((command) =>
        command.kind === 'moveTo' || command.kind === 'lnTo'
          ? [{ x: bounds.x + command.pt.x * sx, y: bounds.y + command.pt.y * sy }]
          : [],
      );
      for (let i = 0; i < points.length; i++)
        if (
          distanceToSegment(point, points[i]!, points[Math.min(i + 1, points.length - 1)]!) <=
          tolerance
        )
          return shape;
    }
  }
  return null;
}

function inside(point: Point, polygon: readonly Point[]): boolean {
  let hit = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const a = polygon[i]!;
    const b = polygon[j]!;
    if (
      a.y > point.y !== b.y > point.y &&
      point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x
    )
      hit = !hit;
  }
  return hit;
}

/** Shapes whose whole bounding box lies inside the lasso, as Lasso Select picks them. */
export function shapesInLasso(slide: SlideData, lasso: readonly Point[]): number[] {
  if (lasso.length < 3) return [];
  return getSlideShapes(slide).flatMap((shape) => {
    const b = getShapeBounds(shape);
    if (!b) return [];
    const corners = [
      { x: b.x, y: b.y },
      { x: b.x + b.w, y: b.y },
      { x: b.x, y: b.y + b.h },
      { x: b.x + b.w, y: b.y + b.h },
    ];
    return corners.every((corner) => inside(corner, lasso)) ? [getShapeId(shape)] : [];
  });
}

// Douglas–Peucker: keeps the corners a hand-drawn outline was meant to have.
function simplify(points: readonly Point[], tolerance: number): Point[] {
  if (points.length < 3) return [...points];
  const first = points[0]!;
  const last = points.at(-1)!;
  let index = 0;
  let farthest = 0;
  for (let i = 1; i < points.length - 1; i++) {
    const distance = distanceToSegment(points[i]!, first, last);
    if (distance > farthest) {
      farthest = distance;
      index = i;
    }
  }
  if (farthest <= tolerance) return [first, last];
  return [
    ...simplify(points.slice(0, index + 1), tolerance).slice(0, -1),
    ...simplify(points.slice(index), tolerance),
  ];
}

// Fractions of the stroke's bounding-box diagonal.
const CLOSED_GAP = 0.2;
const CORNER_TOLERANCE = 0.08;

export type RecognizedShape =
  | { kind: 'line'; from: Point; to: Point }
  | { kind: 'triangle' | 'rect' | 'ellipse'; x: number; y: number; w: number; h: number };

/** What Ink to Shape turns a stroke into, or `null` to keep it as ink. */
export function recognizeShape(points: readonly Point[]): RecognizedShape | null {
  if (points.length < 2) return null;
  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  const w = Math.max(...xs) - x;
  const h = Math.max(...ys) - y;
  const diagonal = Math.hypot(w, h);
  if (diagonal === 0) return null;
  const first = points[0]!;
  const last = points.at(-1)!;
  if (Math.hypot(last.x - first.x, last.y - first.y) > diagonal * CLOSED_GAP) {
    const corners = simplify(points, diagonal * CORNER_TOLERANCE);
    return corners.length === 2 ? { kind: 'line', from: first, to: last } : null;
  }
  // Close the outline, then count its corners.
  const corners = simplify([...points, first], diagonal * CORNER_TOLERANCE).length - 1;
  const kind = corners === 3 ? 'triangle' : corners === 4 ? 'rect' : corners > 4 ? 'ellipse' : null;
  return kind ? { kind, x, y, w, h } : null;
}

/** Adds the recognized shape with the pen's outline and no fill. */
export function addRecognizedShape(
  slide: SlideData,
  shape: RecognizedShape,
  pen: Pen,
): SlideShapeData {
  const width = Math.min(pen.widthEmu, 28575);
  if (shape.kind === 'line')
    return addSlideLine(slide, {
      from: { x: emu(shape.from.x), y: emu(shape.from.y) },
      to: { x: emu(shape.to.x), y: emu(shape.to.y) },
      color: pen.color,
      widthEmu: width,
    });
  const added = addSlideShape(slide, {
    preset: shape.kind,
    x: emu(shape.x),
    y: emu(shape.y),
    w: emu(Math.max(1, shape.w)),
    h: emu(Math.max(1, shape.h)),
  });
  setShapeNoFill(added);
  setShapeStroke(added, { color: pen.color, widthEmu: width, opacity: pen.opacity });
  return added;
}
