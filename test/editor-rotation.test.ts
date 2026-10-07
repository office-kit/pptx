import { describe, expect, it } from 'vitest';
import { selectionBounds, type RotatedRect } from '../packages/editor/src/canvas/rotation.ts';

const shapes: RotatedRect[] = [
  { x: 10, y: 20, w: 80, h: 40, rotation: 0 },
  { x: 140, y: 60, w: 20, h: 80, rotation: 90 },
];

describe('rotated selection bounds', () => {
  it('encloses rotated corners rather than the unrotated bounds', () => {
    expect(selectionBounds(shapes)).toEqual({ x: 10, y: 20, w: 180, h: 90 });
    const diagonal = selectionBounds([{ x: 0, y: 0, w: 100, h: 100, rotation: 45 }]);
    expect(diagonal.w).toBeCloseTo(100 * Math.SQRT2);
    expect(diagonal.h).toBeCloseTo(100 * Math.SQRT2);
    expect(diagonal.x + diagonal.w / 2).toBeCloseTo(50);
    expect(diagonal.y + diagonal.h / 2).toBeCloseTo(50);
  });
});
