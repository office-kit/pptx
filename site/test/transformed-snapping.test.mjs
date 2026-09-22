import assert from 'node:assert/strict';
import test from 'node:test';
import { snapTransformedMove } from '../src/lib/editor/canvas/transformed-snapping.ts';

const slide = { w: 1000, h: 800 };
const rect = { x: 0, y: 0, w: 100, h: 50, rotation: 0 };

test('rotated group children snap to slide edges in slide coordinates', () => {
  const result = snapTransformedMove(
    [rect],
    [],
    [0, 2, -3, 0, 305, 100],
    { x: 0, y: 50 },
    slide,
    6,
  );
  // Visible left is 5 before snapping: moving local y another 5/3 aligns it to 0.
  assert.ok(Math.abs(result.delta.x) < 1e-9);
  assert.ok(Math.abs(result.delta.y - (50 + 5 / 3)) < 1e-9);
  assert.ok(result.guides.some((g) => g.o === 'v' && g.pos === 0));
});

test('snap distance remains in screen axes through reflected nonuniform scaling', () => {
  const matrix = [-4, 0, 0, 2, 800, 100];
  const other = { x: 150, y: 150, w: 100, h: 50, rotation: 0 };
  const near = snapTransformedMove([rect], [other], matrix, { x: 48.75, y: 0 }, slide, 6);
  assert.equal(near.delta.x, 50);
  const far = snapTransformedMove([rect], [other], matrix, { x: 48, y: 0 }, slide, 6);
  assert.equal(far.delta.x, 48);
});

test('rotated child envelopes determine visible edge snapping', () => {
  const result = snapTransformedMove(
    [{ ...rect, x: 20, y: 200, rotation: 90 }],
    [],
    [1, 0, 0, 1, 0, 0],
    { x: -40, y: 0 },
    slide,
    6,
  );
  assert.equal(result.delta.x, -45);
  assert.ok(result.guides.some((g) => g.o === 'v' && g.pos === 0));
});
