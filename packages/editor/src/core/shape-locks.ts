import { getGroupChildren, getShapeId, isShapeLocked, type SlideData } from '@office-kit/pptx';
import { topLevelShapes, selectedShapeIds, type Selection } from './selection.ts';

/** A locked group also prevents gestures on its descendants. */
export function lockedShapeIds(slide: SlideData): Set<number> {
  const locked = new Set<number>();
  const pending = topLevelShapes(slide).map((shape) => ({ shape, inherited: false }));
  while (pending.length) {
    const { shape, inherited } = pending.pop()!;
    const fixed = inherited || isShapeLocked(shape);
    if (fixed) locked.add(getShapeId(shape));
    for (const child of getGroupChildren(shape)) pending.push({ shape: child, inherited: fixed });
  }
  return locked;
}

export function selectionLocked(slide: SlideData | null, selection: Selection): boolean {
  if (!slide) return false;
  const locked = lockedShapeIds(slide);
  return selectedShapeIds(selection).some((id) => locked.has(id));
}
