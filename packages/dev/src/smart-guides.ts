import { slideExtent } from './shape-arrangement.ts';
import type { EditorShape } from './editor.ts';
type Box = NonNullable<EditorShape['bounds']>;
export interface SmartGuide {
  kind?: 'spacing';
  axis: 'x' | 'y';
  value: number;
  from: number;
  to: number;
}
/** Descend through moving objects' ancestors, excluding the moving subtree itself. */
export function smartGuideTargets(roots: EditorShape[], moving: EditorShape[]): EditorShape[] {
  const ids = new Set(moving.map((s) => s.id));
  const containsMoving = (s: EditorShape): boolean =>
    ids.has(s.id) || !!s.children?.some(containsMoving);
  const visit = (s: EditorShape): EditorShape[] => {
    if (s.hidden || ids.has(s.id)) return [];
    return s.children?.some(containsMoving) ? s.children.flatMap(visit) : [s];
  };
  return roots.flatMap(visit);
}
/** Equal edge-to-edge gaps in a row/column; the slide itself is not a spacing object. */
function spacingCandidates(box: Box, targets: Box[], axis: 'x' | 'y', threshold: number) {
  const extent = axis === 'x' ? 'w' : 'h';
  const cross = axis === 'x' ? 'y' : 'x';
  const crossExtent = axis === 'x' ? 'h' : 'w';
  const row = targets
    .filter(
      (target) =>
        Math.min(box[cross] + box[crossExtent], target[cross] + target[crossExtent]) >
        Math.max(box[cross], target[cross]),
    )
    .sort((a, b) => a[axis] - b[axis]);
  const candidates: Array<{ adjustment: number; guides: SmartGuide[] }> = [];
  for (let index = 0; index + 1 < row.length; index++) {
    const first = row[index]!;
    const second = row[index + 1]!;
    const end = first[axis] + first[extent];
    const gap = second[axis] - end;
    if (gap <= 0) continue;
    const low = Math.max(box[cross], first[cross], second[cross]);
    const high = Math.min(
      box[cross] + box[crossExtent],
      first[cross] + first[crossExtent],
      second[cross] + second[crossExtent],
    );
    if (low >= high) continue;
    const positions = [
      end + (gap - box[extent]) / 2,
      first[axis] - gap - box[extent],
      second[axis] + second[extent] + gap,
    ];
    for (const position of positions) {
      const adjustment = position - box[axis];
      if (Math.abs(adjustment) > threshold) continue;
      const finish = position + box[extent];
      if (row.some((target) => position < target[axis] + target[extent] && finish > target[axis]))
        continue;
      const ordered = [first, second, { ...box, [axis]: position }].sort(
        (a, b) => a[axis] - b[axis],
      );
      const gaps = ordered.slice(0, 2).map((target, i) => ({
        kind: 'spacing' as const,
        axis,
        value: (low + high) / 2,
        from: target[axis] + target[extent],
        to: ordered[i + 1]![axis],
      }));
      if (gaps.some((guide) => guide.to <= guide.from)) continue;
      candidates.push({ adjustment, guides: gaps });
    }
  }
  return candidates;
}
export function snapMove(
  moving: EditorShape[],
  others: EditorShape[],
  slide: { width: number; height: number },
  delta: { x: number; y: number },
  threshold: number,
  constrained?: 'x' | 'y',
  options: {
    grid?: { x: number; y: number } | undefined;
    smart: boolean;
    drawing: boolean;
    guides?: Array<{ axis: 'x' | 'y'; offset: number }>;
  } = { smart: true, drawing: false },
): { x: number; y: number; guides: SmartGuide[] } {
  const result = { ...delta, guides: [] as SmartGuide[] };
  const boxes = moving.map(slideExtent).filter((box): box is Box => !!box);
  if (!boxes.length) return result;
  const x = Math.min(...boxes.map((box) => box.x)) + delta.x;
  const y = Math.min(...boxes.map((box) => box.y)) + delta.y;
  const box = {
    x,
    y,
    w: Math.max(...boxes.map((box) => box.x + box.w)) + delta.x - x,
    h: Math.max(...boxes.map((box) => box.y + box.h)) + delta.y - y,
  };
  const ids = new Set(moving.map((shape) => shape.id));
  const targets = (options.smart ? others : [])
    .filter((shape) => !ids.has(shape.id))
    .map(slideExtent)
    .filter((box): box is Box => !!box);
  const spacingTargets = [...targets];
  if (options.smart) targets.push({ x: 0, y: 0, w: slide.width, h: slide.height });
  const alignedTargets: Array<{ guide: SmartGuide; target: Box }> = [];
  for (const axis of ['x', 'y'] as const) {
    if (constrained && constrained !== axis) continue;
    const extent = axis === 'x' ? 'w' : 'h';
    const cross = axis === 'x' ? 'y' : 'x';
    const crossExtent = axis === 'x' ? 'h' : 'w';
    let best: { adjustment: number; guide: SmartGuide; target: Box } | null = null;
    const axisTargets = targets.map((target) => ({ target, fractions: [0, 0.5, 1] }));
    if (options.drawing) {
      const guides = options.guides ?? [
        { axis: 'x', offset: 0 },
        { axis: 'y', offset: 0 },
      ];
      for (const guide of guides) {
        if (guide.axis !== axis) continue;
        const target = { x: 0, y: 0, w: slide.width, h: slide.height };
        target[axis] = target[extent] / 2 + guide.offset;
        target[extent] = 0;
        axisTargets.push({ target, fractions: [0] });
      }
    }
    for (const { target, fractions } of axisTargets) {
      for (const fraction of fractions) {
        const value = target[axis] + target[extent] * fraction;
        for (const sourceFraction of [0, 0.5, 1]) {
          const adjustment = value - box[axis] - box[extent] * sourceFraction;
          if (
            Math.abs(adjustment) > threshold ||
            (best && Math.abs(adjustment) >= Math.abs(best.adjustment))
          )
            continue;
          best = {
            adjustment,
            target,
            guide: {
              axis,
              value,
              from: Math.min(box[cross], target[cross]),
              to: Math.max(box[cross] + box[crossExtent], target[cross] + target[crossExtent]),
            },
          };
        }
      }
    }
    const spacing = options.smart
      ? spacingCandidates(box, spacingTargets, axis, threshold).sort(
          (a, b) => Math.abs(a.adjustment) - Math.abs(b.adjustment),
        )[0]
      : undefined;
    if (spacing && (!best || Math.abs(spacing.adjustment) < Math.abs(best.adjustment))) {
      result[axis] += spacing.adjustment;
      result.guides.push(...spacing.guides);
      continue;
    }
    if (!best && options.grid && Number.isFinite(options.grid[axis]) && options.grid[axis] > 0) {
      const spacing = options.grid[axis];
      result[axis] += Math.round(box[axis] / spacing) * spacing - box[axis];
    }
    if (best) {
      if (spacing && Math.abs(spacing.adjustment - best.adjustment) < 0.001)
        result.guides.push(...spacing.guides);
      result[axis] += best.adjustment;
      result.guides.push(best.guide);
      alignedTargets.push({ guide: best.guide, target: best.target });
    }
  }
  // Both axis corrections must be known before spanning the final selection bounds.
  alignedTargets.forEach(({ guide, target }) => {
    const cross = guide.axis === 'x' ? 'y' : 'x';
    const extent = guide.axis === 'x' ? 'h' : 'w';
    const start = box[cross] + result[cross] - delta[cross];
    guide.from = Math.min(start, target[cross]);
    guide.to = Math.max(start + box[extent], target[cross] + target[extent]);
  });
  return result;
}
