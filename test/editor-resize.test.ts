import { describe, expect, it } from 'vitest';
import {
  resizeRect,
  resizeSelectionRects,
  type ResizeHandle,
} from '../site/src/lib/editor/canvas/resize.ts';
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
const quarterTurnHandles: Record<ResizeHandle, ResizeHandle> = {
  nw: 'sw',
  n: 'w',
  ne: 'nw',
  e: 'n',
  se: 'ne',
  s: 'e',
  sw: 'se',
  w: 's',
};
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

describe('individual selection resize', () => {
  it('uses the grabbed shape ratios without moving the other objects', () => {
    const shapes = [
      { x: 362, y: 609, w: 1046, h: 148 },
      { x: 420, y: 1155, w: 582, h: 118 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -100, y: 40 },
      { w: 1, h: 1 },
    );
    expect(result[0]).toEqual({ x: 362, y: 609, w: 946, h: 188 });
    expect(result[1]!.x).toBe(420);
    expect(result[1]!.y).toBe(1155);
    expect(result[1]!.w).toBeCloseTo((582 * 946) / 1046);
    expect(result[1]!.h).toBeCloseTo((118 * 188) / 148);
  });
  for (const handle of handles) {
    it(`keeps each opposite ${handle} anchor when resizing rotated selections`, () => {
      const shapes = [
        { ...original, rotation: 30 },
        { x: 700, y: 500, w: 150, h: 80, rotation: 120 },
      ];
      const result = resizeSelectionRects(
        shapes,
        shapes[0]!,
        handle,
        { x: 30, y: 20 },
        { w: 1, h: 1 },
      );
      result.forEach((rect, i) => {
        const shape = shapes[i]!;
        const targetHandle = i === 0 ? handle : quarterTurnHandles[handle];
        const before = point(shape, targetHandle, shape.rotation, true);
        const after = point(rect, targetHandle, shape.rotation, true);
        expect(after.x).toBeCloseTo(before.x);
        expect(after.y).toBeCloseTo(before.y);
        const swapsAxes = i === 1;
        expect(rect.w / shape.w).toBeCloseTo(
          swapsAxes ? result[0]!.h / shapes[0]!.h : result[0]!.w / shapes[0]!.w,
        );
        expect(rect.h / shape.h).toBeCloseTo(
          swapsAxes ? result[0]!.w / shapes[0]!.w : result[0]!.h / shapes[0]!.h,
        );
      });
    });
  }
  it('matches independent local-axis scaling after both objects rotate 90 degrees', () => {
    const shapes = [
      { x: 362, y: 609, w: 1046, h: 148, rotation: 90 },
      { x: 420, y: 1155, w: 582, h: 118, rotation: 90 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'ne',
      { x: 40, y: -100 },
      { w: 1, h: 1 },
    );
    expect(result[0]!.w).toBeCloseTo(946);
    expect(result[0]!.h).toBeCloseTo(188);
    expect(result[1]!.w).toBeCloseTo((582 * 946) / 1046);
    expect(result[1]!.h).toBeCloseTo((118 * 188) / 148);
    for (const [i, shape] of shapes.entries()) {
      const before = point(shape, 'ne', 90, true);
      const after = point(result[i]!, 'ne', 90, true);
      expect(after.x).toBeCloseTo(before.x);
      expect(after.y).toBeCloseTo(before.y);
    }
  });
  it('maps a screen south-east handle to the target local north-east handle at 90 degrees', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      { x: 500, y: 300, w: 200, h: 80, rotation: 90 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -28.8, y: 28.8 },
      { w: 1, h: 1 },
    );
    // The 0-degree title establishes screen scaleX=.904, scaleY=1.288.
    // For the 90-degree target those screen axes are its local height/width.
    expect(result[1]!.w).toBeCloseTo(200 * 1.288);
    expect(result[1]!.h).toBeCloseTo(80 * 0.904);
    const before = point(shapes[1]!, 'ne', 90, true);
    const after = point(result[1]!, 'ne', 90, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  it('maps a south-east handle at a positive 45-degree target', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      { x: 500, y: 300, w: 200, h: 80, rotation: 45 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -28.8, y: 28.8 },
      { w: 1, h: 1 },
    );
    expect(result[1]!.w).toBeCloseTo(200 * 1.288);
    expect(result[1]!.h).toBeCloseTo(80 * 0.904);
    const before = point(shapes[1]!, 'ne', 45, true);
    const after = point(result[1]!, 'ne', 45, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  it('maps a south-east handle at a negative 45-degree target', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      { x: 500, y: 300, w: 200, h: 80, rotation: -45 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -28.8, y: 28.8 },
      { w: 1, h: 1 },
    );
    expect(result[1]!.w).toBeCloseTo(200 * 1.288);
    expect(result[1]!.h).toBeCloseTo(80 * 0.904);
    const before = point(shapes[1]!, 'sw', -45, true);
    const after = point(result[1]!, 'sw', -45, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  it('treats a 315-degree target as equivalent to negative 45 degrees', () => {
    const target = { x: 500, y: 300, w: 200, h: 80 };
    const delta = { x: -28.8, y: 28.8 };
    const negative = resizeSelectionRects(
      [
        { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
        { ...target, rotation: -45 },
      ],
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      'se',
      delta,
      { w: 1, h: 1 },
    )[1]!;
    const normalized = resizeSelectionRects(
      [
        { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
        { ...target, rotation: 315 },
      ],
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      'se',
      delta,
      { w: 1, h: 1 },
    )[1]!;
    expect(normalized.x).toBeCloseTo(negative.x);
    expect(normalized.y).toBeCloseTo(negative.y);
    expect(normalized.w).toBeCloseTo(negative.w);
    expect(normalized.h).toBeCloseTo(negative.h);
  });
  it('keeps the local handle for a 30-to-60 degree relative rotation', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 30 },
      { x: 500, y: 300, w: 200, h: 80, rotation: 60 },
    ];
    const angle = (30 * Math.PI) / 180;
    const delta = {
      x: -30 * Math.cos(angle) - 20 * Math.sin(angle),
      y: -30 * Math.sin(angle) + 20 * Math.cos(angle),
    };
    const result = resizeSelectionRects(shapes, shapes[0]!, 'se', delta, { w: 1, h: 1 });
    expect(result[1]!.w).toBeCloseTo(180);
    expect(result[1]!.h).toBeCloseTo(96);
    const before = point(shapes[1]!, 'se', 60, true);
    const after = point(result[1]!, 'se', 60, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  it('maps the opposite quarter-turn for a negative 90 degree target', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      { x: 500, y: 300, w: 200, h: 80, rotation: -90 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -28.8, y: 28.8 },
      { w: 1, h: 1 },
    );
    expect(result[1]!.w).toBeCloseTo(200 * 1.288);
    expect(result[1]!.h).toBeCloseTo(80 * 0.904);
    const before = point(shapes[1]!, 'sw', -90, true);
    const after = point(result[1]!, 'sw', -90, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  it('preserves axes and flips the local handle at a 180 degree target', () => {
    const shapes = [
      { x: 100, y: 200, w: 300, h: 100, rotation: 0 },
      { x: 500, y: 300, w: 200, h: 80, rotation: 180 },
    ];
    const result = resizeSelectionRects(
      shapes,
      shapes[0]!,
      'se',
      { x: -28.8, y: 28.8 },
      { w: 1, h: 1 },
    );
    expect(result[1]!.w).toBeCloseTo(200 * 0.904);
    expect(result[1]!.h).toBeCloseTo(80 * 1.288);
    const before = point(shapes[1]!, 'nw', 180, true);
    const after = point(result[1]!, 'nw', 180, true);
    expect(after.x).toBeCloseTo(before.x);
    expect(after.y).toBeCloseTo(before.y);
  });
  for (const rotation of [135, -135]) {
    it(`flips the local handle without swapping axes at ${rotation}-degree target`, () => {
      const shapes = [
        { x: 362, y: 609, w: 1046, h: 148, rotation: 0 },
        { x: 420, y: 1155, w: 582, h: 118, rotation },
      ];
      const result = resizeSelectionRects(
        shapes,
        shapes[0]!,
        'se',
        { x: -100, y: 40 },
        { w: 1, h: 1 },
      );
      const widthScale = result[0]!.w / shapes[0]!.w;
      const heightScale = result[0]!.h / shapes[0]!.h;
      expect(result[1]!.w).toBeCloseTo(shapes[1]!.w * widthScale);
      expect(result[1]!.h).toBeCloseTo(shapes[1]!.h * heightScale);
      // The target local north-west handle is dragged; its opposite south-east stays fixed.
      const before = point(shapes[1]!, 'nw', rotation, true);
      const after = point(result[1]!, 'nw', rotation, true);
      expect(after.x).toBeCloseTo(before.x);
      expect(after.y).toBeCloseTo(before.y);
    });
  }
  it('keeps zero-height lines flat and does not translate between objects', () => {
    const shapes = [
      { x: 0, y: 10, w: 100, h: 0 },
      { x: 150, y: 10, w: 50, h: 0 },
    ];
    expect(
      resizeSelectionRects(shapes, shapes[0]!, 'e', { x: 100, y: 60 }, { w: 10, h: 10 }),
    ).toEqual([
      { x: 0, y: 10, w: 200, h: 0 },
      { x: 150, y: 10, w: 100, h: 0 },
    ]);
  });
});
