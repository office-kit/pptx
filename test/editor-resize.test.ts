import { describe, expect, it } from 'vitest';
import {
  resizeRect,
  resizeSelectionRects,
  type ResizeHandle,
} from '../site/src/lib/editor/canvas/resize.ts';
import { selectionBounds } from '../site/src/lib/editor/canvas/rotation.ts';
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

describe('proportional selection resize', () => {
  for (const handle of ['nw', 'ne', 'se', 'sw'] as const) {
    for (const rotation of [0, 45, 90, 315]) {
      it(`scales ${handle} with a ${rotation}° object about the opposite corner`, () => {
        const shapes = [
          { x: 10, y: 20, w: 100, h: 40, rotation: 0 },
          { x: 200, y: 100, w: 80, h: 20, rotation },
        ];
        const frame = selectionBounds(shapes);
        const result = resizeSelectionRects(
          shapes,
          frame,
          handle,
          {
            x: ((handle.includes('w') ? -1 : 1) * frame.w) / 2,
            y: ((handle.includes('n') ? -1 : 1) * frame.h) / 2,
          },
          { w: 1, h: 1 },
        );
        const anchor = point(frame, handle, 0, true);
        const resizedFrame = selectionBounds(
          result.map((rect, i) => ({ ...rect, rotation: shapes[i]!.rotation })),
        );
        const after = point(resizedFrame, handle, 0, true);
        expect(after.x).toBeCloseTo(anchor.x);
        expect(after.y).toBeCloseTo(anchor.y);
        expect(resizedFrame.w).toBeCloseTo(frame.w * 1.5);
        expect(resizedFrame.h).toBeCloseTo(frame.h * 1.5);
        result.forEach((rect, i) => {
          const start = shapes[i]!;
          expect(rect.w).toBeCloseTo(start.w * 1.5);
          expect(rect.h).toBeCloseTo(start.h * 1.5);
          expect(rect.x).toBeCloseTo(anchor.x + (start.x - anchor.x) * 1.5);
          expect(rect.y).toBeCloseTo(anchor.y + (start.y - anchor.y) * 1.5);
        });
      });
    }
  }
  it('keeps a fixed corner and relative layout when clamping a selection to minimum size', () => {
    expect(
      resizeSelectionRects(
        [
          { x: 0, y: 0, w: 100, h: 100 },
          { x: 100, y: 0, w: 100, h: 100 },
        ],
        { x: 0, y: 0, w: 200, h: 100 },
        'nw',
        { x: 500, y: 500 },
        { w: 20, h: 20 },
      ),
    ).toEqual([
      { x: 160, y: 80, w: 20, h: 20 },
      { x: 180, y: 80, w: 20, h: 20 },
    ]);
  });
  it('scales zero-height lines without giving them thickness', () => {
    expect(
      resizeSelectionRects(
        [
          { x: 0, y: 10, w: 100, h: 0 },
          { x: 150, y: 10, w: 50, h: 0 },
        ],
        { x: 0, y: 10, w: 200, h: 0 },
        'se',
        { x: 200, y: 60 },
        { w: 10, h: 10 },
      ),
    ).toEqual([
      { x: 0, y: 10, w: 200, h: 0 },
      { x: 300, y: 10, w: 100, h: 0 },
    ]);
  });
  it('scales zero-width lines using their height', () => {
    expect(
      resizeSelectionRects(
        [
          { x: 10, y: 0, w: 0, h: 100 },
          { x: 10, y: 150, w: 0, h: 50 },
        ],
        { x: 10, y: 0, w: 0, h: 200 },
        'nw',
        { x: -60, y: -200 },
        { w: 10, h: 10 },
      ),
    ).toEqual([
      { x: 10, y: -200, w: 0, h: 200 },
      { x: 10, y: 100, w: 0, h: 100 },
    ]);
  });
  it('leaves a selection of coincident zero-size points finite', () => {
    const point = { x: 10, y: 20, w: 0, h: 0 };
    expect(
      resizeSelectionRects([point, point], point, 'se', { x: 20, y: 30 }, { w: 10, h: 10 }),
    ).toEqual([point, point]);
  });
});
