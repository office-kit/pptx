import { test } from 'node:test';
import assert from 'node:assert/strict';
import { dragPictureCrop } from '../src/picture-crop.ts';
const shape = {
  bounds: { x: 100000, y: 200000, w: 4000000, h: 2000000 },
  rotation: 0,
  imageCrop: null,
  flip: null,
};
test('crop edge changes frame without resizing source image', () => {
  const result = dragPictureCrop(shape, 'e', -1000000, 0);
  assert.deepEqual(result.bounds, { ...shape.bounds, w: 3000000 });
  assert.deepEqual(result.imageCrop, { left: 0, right: 0.25, top: 0, bottom: 0 });
  const corner = dragPictureCrop(shape, 'nw', 1000000, 500000);
  assert.deepEqual(corner.bounds, { x: 1100000, y: 700000, w: 3000000, h: 1500000 });
  assert.deepEqual(corner.imageCrop, { left: 0.25, right: 0, top: 0.25, bottom: 0 });
});
test('source dragging preserves frame and respects flips', () => {
  const result = dragPictureCrop(shape, '', 1000000, 500000);
  assert.deepEqual(result.bounds, shape.bounds);
  assert.deepEqual(result.imageCrop, { left: -0.25, right: 0.25, top: -0.25, bottom: 0.25 });
  const flipped = dragPictureCrop(
    { ...shape, flip: { horizontal: true, vertical: true } },
    '',
    1000000,
    500000,
  );
  assert.deepEqual(flipped.imageCrop, { left: 0.25, right: -0.25, top: 0.25, bottom: -0.25 });
});
test('rotated and flipped edge crop keeps opposite edge fixed', () => {
  const result = dragPictureCrop(
    { ...shape, rotation: 90, flip: { horizontal: true } },
    'e',
    0,
    -1000000,
  );
  assert.deepEqual(result.bounds, { x: 600000, y: -300000, w: 3000000, h: 2000000 });
  assert.deepEqual(result.imageCrop, { left: 0.25, right: 0, top: 0, bottom: 0 });
});
test('existing signed crop and outward drag preserve image scale', () => {
  const result = dragPictureCrop(
    { ...shape, imageCrop: { left: -0.5, right: -0.5, top: 0.1, bottom: 0.2 } },
    'w',
    -1000000,
    0,
  );
  assert.equal(result.bounds.w, 5000000);
  assert.equal(result.imageCrop.left, -1);
  assert.equal(result.imageCrop.right, -0.5);
  assert.equal(result.imageCrop.top, 0.1);
  assert.ok(dragPictureCrop(shape, 'e', -1e8, 0).bounds.w > 0);
});

test('ratio-constrained corners and sides preserve source scale and opposite anchors', () => {
  for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
    const result = dragPictureCrop(shape, handle, 300000, 500000, 2);
    assert.equal(result.bounds.w / result.bounds.h, 2);
    const b = result.bounds,
      c = result.imageCrop;
    assert.ok(Math.abs(b.w / (1 - c.left - c.right) - shape.bounds.w) < 0.001);
    assert.ok(Math.abs(b.h / (1 - c.top - c.bottom) - shape.bounds.h) < 0.001);
    if (handle.includes('w')) assert.equal(b.x + b.w, shape.bounds.x + shape.bounds.w);
    if (handle.includes('e')) assert.equal(b.x, shape.bounds.x);
    if (handle.includes('n')) assert.equal(b.y + b.h, shape.bounds.y + shape.bounds.h);
    if (handle.includes('s')) assert.equal(b.y, shape.bounds.y);
    if (handle.length === 1 && ['e', 'w'].includes(handle))
      assert.equal(b.y + b.h / 2, shape.bounds.y + shape.bounds.h / 2);
  }
  const vertical = dragPictureCrop(shape, 'se', 0, -500000, 2);
  assert.equal(vertical.bounds.w, 3000000);
  assert.equal(vertical.bounds.h, 1500000);
  const crossed = dragPictureCrop(shape, 'se', -1e8, -1e8, 2);
  assert.equal(crossed.bounds.w / crossed.bounds.h, 2);
  assert.ok(crossed.bounds.h > 0);
});

test('source corner resizing fixes the crop frame and opposite source corner across flips and rotation', () => {
  for (const horizontal of [false, true])
    for (const vertical of [false, true]) {
      const input = {
        ...shape,
        rotation: 90,
        flip: { horizontal, vertical },
        imageCrop: { left: 0.1, right: 0.2, top: -0.1, bottom: 0.2 },
      };
      const sourceRect = (crop) => {
        const w = shape.bounds.w / (1 - crop.left - crop.right),
          h = shape.bounds.h / (1 - crop.top - crop.bottom);
        return {
          w,
          h,
          x: -(horizontal ? crop.right : crop.left) * w,
          y: -(vertical ? crop.bottom : crop.top) * h,
        };
      };
      const before = sourceRect(input.imageCrop);
      for (const corner of ['nw', 'ne', 'sw', 'se']) {
        const localX = corner.includes('w') ? -400000 : 400000;
        const localY = corner.includes('n') ? -200000 : 200000;
        const result = dragPictureCrop(input, `source-${corner}`, -localY, localX);
        assert.deepEqual(result.bounds, input.bounds);
        const after = sourceRect(result.imageCrop);
        assert.ok(after.w > before.w);
        assert.ok(Math.abs(after.w / after.h - before.w / before.h) < 1e-10);
        assert.ok(
          Math.abs(
            (corner.includes('w') ? after.x + after.w : after.x) -
              (corner.includes('w') ? before.x + before.w : before.x),
          ) < 0.001,
        );
        assert.ok(
          Math.abs(
            (corner.includes('n') ? after.y + after.h : after.y) -
              (corner.includes('n') ? before.y + before.h : before.y),
          ) < 0.001,
        );
      }
    }
  const crossed = dragPictureCrop(shape, 'source-se', -1e9, -1e9);
  assert.ok(shape.bounds.w / (1 - crossed.imageCrop.left - crossed.imageCrop.right) > 0);
});
