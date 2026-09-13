// Smart-guide snapping — the behaviour that makes dragging feel precise.
//
// When a shape's edge or centre lines up (within a small threshold) with
// another shape's edge/centre or the slide's edges/centre, we nudge it into
// exact alignment and report a guide line to draw. All coordinates are EMU.

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Guide {
  /** 'v' = vertical line at x=pos; 'h' = horizontal line at y=pos. */
  readonly o: 'v' | 'h';
  readonly pos: number;
  /** Span on the perpendicular axis (for drawing a tidy segment). */
  readonly from: number;
  readonly to: number;
}

export interface SnapResult {
  readonly x: number;
  readonly y: number;
  readonly guides: readonly Guide[];
}

interface Anchor {
  readonly value: number;
  readonly kind: 'start' | 'center' | 'end';
}

function xAnchors(r: Rect): Anchor[] {
  return [
    { value: r.x, kind: 'start' },
    { value: r.x + r.w / 2, kind: 'center' },
    { value: r.x + r.w, kind: 'end' },
  ];
}
function yAnchors(r: Rect): Anchor[] {
  return [
    { value: r.y, kind: 'start' },
    { value: r.y + r.h / 2, kind: 'center' },
    { value: r.y + r.h, kind: 'end' },
  ];
}

/**
 * Snap `moving` against `others` + the slide box. Returns the adjusted x/y and
 * the guide lines to render. `thresh` is the snap distance in EMU.
 */
export function snapMove(
  moving: Rect,
  others: readonly Rect[],
  slide: { w: number; h: number },
  thresh: number,
): SnapResult {
  const targets: Rect[] = [{ x: 0, y: 0, w: slide.w, h: slide.h }, ...others];

  // Best snap per axis (smallest delta wins).
  let bestX: { delta: number; pos: number; targets: Rect[] } | null = null;
  let bestY: { delta: number; pos: number; targets: Rect[] } | null = null;

  for (const ma of xAnchors(moving)) {
    for (const t of targets) {
      for (const ta of xAnchors(t)) {
        const delta = ta.value - ma.value;
        if (Math.abs(delta) <= thresh && (!bestX || Math.abs(delta) < Math.abs(bestX.delta))) {
          bestX = { delta, pos: ta.value, targets: [t] };
        } else if (bestX && ta.value === bestX.pos) {
          bestX.targets.push(t);
        }
      }
    }
  }
  for (const ma of yAnchors(moving)) {
    for (const t of targets) {
      for (const ta of yAnchors(t)) {
        const delta = ta.value - ma.value;
        if (Math.abs(delta) <= thresh && (!bestY || Math.abs(delta) < Math.abs(bestY.delta))) {
          bestY = { delta, pos: ta.value, targets: [t] };
        } else if (bestY && ta.value === bestY.pos) {
          bestY.targets.push(t);
        }
      }
    }
  }

  const x = moving.x + (bestX?.delta ?? 0);
  const y = moving.y + (bestY?.delta ?? 0);
  const snapped: Rect = { ...moving, x, y };

  const guides: Guide[] = [];
  if (bestX) {
    const ys = [
      snapped.y,
      snapped.y + snapped.h,
      ...bestX.targets.flatMap((t) => [t.y, t.y + t.h]),
    ];
    guides.push({ o: 'v', pos: bestX.pos, from: Math.min(...ys), to: Math.max(...ys) });
  }
  if (bestY) {
    const xs = [
      snapped.x,
      snapped.x + snapped.w,
      ...bestY.targets.flatMap((t) => [t.x, t.x + t.w]),
    ];
    guides.push({ o: 'h', pos: bestY.pos, from: Math.min(...xs), to: Math.max(...xs) });
  }

  return { x, y, guides };
}
