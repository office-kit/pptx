import { describe, expect, it } from 'vitest';
import {
  rotateRect,
  selectionBounds,
  type RotatedRect,
} from '../site/src/lib/editor/canvas/rotation.ts';

const shapes: RotatedRect[] = [
  { x: 10, y: 20, w: 80, h: 40, rotation: 0 },
  { x: 140, y: 60, w: 20, h: 80, rotation: 90 },
];

describe('selection rotation', () => {
  it('encloses rotated corners rather than the unrotated bounds', () => {
    expect(selectionBounds(shapes)).toEqual({ x: 10, y: 20, w: 180, h: 90 });
    const diagonal = selectionBounds([{ x: 0, y: 0, w: 100, h: 100, rotation: 45 }]);
    expect(diagonal.w).toBeCloseTo(100 * Math.SQRT2);
    expect(diagonal.h).toBeCloseTo(100 * Math.SQRT2);
    expect(diagonal.x + diagonal.w / 2).toBeCloseTo(50);
    expect(diagonal.y + diagonal.h / 2).toBeCloseTo(50);
  });
  for (const degrees of [-450, -90, -15, 0, 30, 90, 180, 360, 450]) {
    it(`preserves sizes, distance and relative angles through ${degrees}°`, () => {
      const pivot = { x: 100, y: 65 };
      const result = shapes.map((shape) => rotateRect(shape, pivot, degrees));
      const centre = (rect: RotatedRect) => ({ x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 });
      for (let i = 0; i < shapes.length; i++) {
        const original = shapes[i]!;
        const rotated = result[i]!;
        expect(rotated.w).toBe(original.w);
        expect(rotated.h).toBe(original.h);
        expect(rotated.rotation).toBe((((original.rotation + degrees) % 360) + 360) % 360);
        const a = centre(original),
          b = centre(rotated);
        expect(Math.hypot(b.x - pivot.x, b.y - pivot.y)).toBeCloseTo(
          Math.hypot(a.x - pivot.x, a.y - pivot.y),
        );
        const restored = rotateRect(rotated, pivot, -degrees);
        expect(restored.x).toBeCloseTo(original.x);
        expect(restored.y).toBeCloseTo(original.y);
        expect(restored.rotation).toBe(original.rotation);
      }
      const a = centre(result[0]!),
        b = centre(result[1]!);
      expect(Math.hypot(b.x - a.x, b.y - a.y)).toBeCloseTo(Math.hypot(100, 60));
    });
  }
  it('rotates the centres by a quarter turn around a shared pivot', () => {
    const result = rotateRect(shapes[0]!, { x: 100, y: 65 }, 90);
    expect(result.x).toBeCloseTo(85);
    expect(result.y).toBeCloseTo(-5);
    expect(result.rotation).toBe(90);
  });
});
