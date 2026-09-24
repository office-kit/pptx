import { test } from 'node:test';
import assert from 'node:assert/strict';
import { alignedPositions, distributedPositions, slideExtent } from '../src/shape-arrangement.ts';

const shape = (id, x, y, w, h, extra = {}) => ({
  id,
  bounds: { x, y, w, h },
  rotation: 0,
  flip: null,
  ...extra,
});
const close = (a, b, tolerance = 4) => assert.ok(Math.abs(a - b) < tolerance, `${a} != ${b}`);
const apply = (shapes, positions) =>
  shapes.map((s) => ({ ...s, bounds: positions.find((p) => p.id === s.id)?.bounds ?? s.bounds }));
const shapes = [
  shape(1, 100, 200, 180, 80, { rotation: 30, parentTransform: [0, 2, -3, 0, 2000, 500] }),
  shape(2, 800, 500, 250, 140, { rotation: 55, parentTransform: [-2, 0, 0, 3, 3000, 400] }),
  shape(3, 1200, 300, 120, 250, { rotation: 10 }),
];
test('slide extent projects rotation, reflection and parent scale', () => {
  const box = slideExtent(
    shape(1, 10, 20, 100, 50, {
      rotation: 90,
      flip: { horizontal: true, vertical: false },
      parentTransform: [0, 2, -3, 0, 1000, 2000],
    }),
  );
  for (const [key, expected] of Object.entries({ x: 715, y: 2070, w: 300, h: 100 }))
    close(box[key], expected, 1e-8);
  assert.equal(slideExtent(shape(1, 0, 0, 10, 10, { parentTransform: [0, 0, 0, 1, 0, 0] })), null);
});
test('all six alignments use slide-space extents and retain local sizes and orientation', () => {
  const original = structuredClone(shapes);
  const size = { width: 10000, height: 7000 };
  for (let direction = 0; direction < 6; direction++) {
    const horizontal = direction < 3,
      axis = horizontal ? 'x' : 'y',
      extent = horizontal ? 'w' : 'h';
    const anchor = (b) => b[axis] + ((direction % 3) * b[extent]) / 2;
    const before = shapes.map(slideExtent);
    const expected =
      direction % 3 === 0
        ? Math.min(...before.map((b) => b[axis]))
        : direction % 3 === 2
          ? Math.max(...before.map((b) => b[axis] + b[extent]))
          : (Math.min(...before.map((b) => b[axis])) +
              Math.max(...before.map((b) => b[axis] + b[extent]))) /
            2;
    const after = apply(shapes, alignedPositions(shapes, direction, size));
    after.forEach((s, i) => {
      close(anchor(slideExtent(s)), expected);
      close(slideExtent(s)[horizontal ? 'y' : 'x'], before[i][horizontal ? 'y' : 'x']);
      assert.equal(s.bounds.w, shapes[i].bounds.w);
      assert.equal(s.bounds.h, shapes[i].bounds.h);
      assert.equal(s.rotation, shapes[i].rotation);
    });
    const single = apply([shapes[0]], alignedPositions([shapes[0]], direction, size))[0];
    close(
      anchor(slideExtent(single)),
      ((direction % 3) * (horizontal ? size.width : size.height)) / 2,
    );
  }
  assert.deepEqual(shapes, original);
});
test('distribution keeps endpoints fixed and equalizes visual gaps across coordinate spaces', () => {
  for (const axis of ['x', 'y']) {
    const extent = axis === 'x' ? 'w' : 'h';
    const ordered = [...shapes].sort((a, b) => slideExtent(a)[axis] - slideExtent(b)[axis]);
    const positions = distributedPositions(ordered.toReversed(), axis);
    const after = apply(ordered, positions).map(slideExtent);
    const before = ordered.map(slideExtent);
    close(after[0][axis], before[0][axis]);
    close(after.at(-1)[axis], before.at(-1)[axis]);
    close(
      after[1][axis] - after[0][axis] - after[0][extent],
      after[2][axis] - after[1][axis] - after[1][extent],
    );
    after.forEach((b, i) =>
      close(b[axis === 'x' ? 'y' : 'x'], before[i][axis === 'x' ? 'y' : 'x']),
    );
  }
  assert.deepEqual(distributedPositions(shapes.slice(0, 2), 'x'), []);
});

test('explicit slide alignment centers multiple objects independently', () => {
  const size = { width: 10000, height: 7000 };
  for (const direction of [1, 4]) {
    const after = apply(shapes, alignedPositions(shapes, direction, size, true));
    for (const s of after) {
      const box = slideExtent(s);
      close(
        direction === 1 ? box.x + box.w / 2 : box.y + box.h / 2,
        direction === 1 ? size.width / 2 : size.height / 2,
      );
    }
  }
});
test('slide distribution includes equal margins and supports two objects', () => {
  const size = { width: 10000, height: 7000 };
  for (const axis of ['x', 'y']) {
    const extent = axis === 'x' ? 'w' : 'h';
    for (const count of [2, 3]) {
      const subset = shapes.slice(0, count);
      const after = apply(subset, distributedPositions(subset, axis, size))
        .map(slideExtent)
        .sort((a, b) => a[axis] - b[axis]);
      const gap = after[0][axis];
      for (let i = 1; i < after.length; i++)
        close(after[i][axis] - after[i - 1][axis] - after[i - 1][extent], gap);
      close(
        (axis === 'x' ? size.width : size.height) - after.at(-1)[axis] - after.at(-1)[extent],
        gap,
      );
    }
  }
});
