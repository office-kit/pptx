import { describe, expect, it } from 'vitest';
import { resizeRect, type ResizeHandle } from '../site/src/lib/editor/canvas/resize.ts';
import type { Rect } from '../site/src/lib/editor/canvas/snapping.ts';

const handles: ResizeHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const original = { x: 100, y: 200, w: 300, h: 100 };
function point(rect: Rect, handle: ResizeHandle, rotation: number, opposite = false) {
  const direction = opposite ? -1 : 1;
  const x = (((handle.includes('e') ? 1 : handle.includes('w') ? -1 : 0) * rect.w) / 2) * direction;
  const y = (((handle.includes('s') ? 1 : handle.includes('n') ? -1 : 0) * rect.h) / 2) * direction;
  const angle = (rotation * Math.PI) / 180;
  return {
    x: rect.x + rect.w / 2 + x * Math.cos(angle) - y * Math.sin(angle),
    y: rect.y + rect.h / 2 + x * Math.sin(angle) + y * Math.cos(angle),
  };
}
describe('canvas resize geometry', () => {
  for (const rotation of [0, 45, 90, 180, 315]) {
    for (const handle of handles) {
      it(`keeps the opposite ${handle} handle fixed at ${rotation} degrees`, () => {
        const result = resizeRect(
          original,
          handle,
          { x: 30, y: 20 },
          rotation,
          { w: 10, h: 10 },
          false,
        );
        const before = point(original, handle, rotation, true);
        const after = point(result, handle, rotation, true);
        expect(after.x).toBeCloseTo(before.x, 8);
        expect(after.y).toBeCloseTo(before.y, 8);
        if (handle.length === 2) {
          const movingBefore = point(original, handle, rotation);
          const movingAfter = point(result, handle, rotation);
          expect(movingAfter.x).toBeCloseTo(movingBefore.x + 30, 8);
          expect(movingAfter.y).toBeCloseTo(movingBefore.y + 20, 8);
        }
      });
    }
  }
  for (const handle of handles) {
    it(`preserves aspect and its fixed anchor for Shift-${handle}`, () => {
      const result = resizeRect(original, handle, { x: 70, y: -45 }, 35, { w: 10, h: 10 }, true);
      expect(result.w / result.h).toBeCloseTo(3, 8);
      const before = point(original, handle, 35, true);
      const after = point(result, handle, 35, true);
      expect(after.x).toBeCloseTo(before.x, 8);
      expect(after.y).toBeCloseTo(before.y, 8);
    });
  }
  it('clamps north-west shrinking without moving the opposite corner', () => {
    expect(resizeRect(original, 'nw', { x: 1000, y: 1000 }, 0, { w: 10, h: 20 }, false)).toEqual({
      x: 390,
      y: 280,
      w: 10,
      h: 20,
    });
    const aspect = resizeRect(original, 'nw', { x: 1000, y: 1000 }, 0, { w: 10, h: 20 }, true);
    expect(aspect).toEqual({ x: 340, y: 280, w: 60, h: 20 });
  });
  it('does not enlarge an untouched dimension of a thin line', () => {
    expect(
      resizeRect({ x: 0, y: 0, w: 100, h: 0 }, 'e', { x: 20, y: 80 }, 0, { w: 10, h: 10 }, false),
    ).toEqual({ x: 0, y: 0, w: 120, h: 0 });
  });
});
