import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  groupShapes,
  setShapeBounds,
  setShapeRotation,
  setShapeFlip,
  getShapeId,
  getGroupChildren,
  getSlides,
  savePresentation,
  loadPresentation,
  emu,
} from '@office-kit/pptx';
import {
  shapeScope,
  project,
  invert,
  compose,
  IDENTITY,
} from '../src/lib/editor/canvas/group-space.ts';

const near = (actual, expected) => {
  assert.ok(Math.abs(actual.x - expected.x) < 1e-7, `${actual.x} ≠ ${expected.x}`);
  assert.ok(Math.abs(actual.y - expected.y) < 1e-7, `${actual.y} ≠ ${expected.y}`);
};
const box = (slide, x, y, w, h, text) =>
  addSlideTextBox(slide, { x: emu(x), y: emu(y), w: emu(w), h: emu(h), text });

test('nested, anisotropically scaled and reflected groups project points and pointer deltas after reload', async () => {
  const pres = createPresentation(),
    slide = addBlankSlide(pres);
  const a = box(slide, 10, 20, 20, 10, '日本語'),
    b = box(slide, 50, 40, 10, 20, 'English');
  const inner = groupShapes([a, b]);
  setShapeBounds(inner, { x: emu(100), y: emu(200), w: emu(100), h: emu(120) });
  setShapeRotation(inner, 90);
  setShapeFlip(inner, { horizontal: true, vertical: false });
  const sibling = box(slide, 300, 300, 20, 20, 'Outside inner');
  const outer = groupShapes([inner, sibling]);
  setShapeBounds(outer, { x: emu(500), y: emu(600), w: emu(440), h: emu(60) });
  setShapeRotation(outer, 90);
  const loaded = getSlides(await loadPresentation(await savePresentation(pres)))[0];
  const scope = shapeScope(loaded, getShapeId(a));
  assert.equal(getShapeId(scope.parent), getShapeId(inner));
  assert.deepEqual(scope.textScale, { x: 4, y: 1.5 });
  assert.deepEqual(scope.shapes.map(getShapeId), [getShapeId(a), getShapeId(b)]);
  // Child point (10,20): inner scale+flip+90° → (210,310).
  // Outer scale → (720,655), rotated 90° around (720,630) → (695,630).
  near(project(scope.matrix, { x: 10, y: 20 }), { x: 695, y: 630 });
  const inverse = invert(scope.matrix);
  assert.ok(inverse);
  const start = { x: 23, y: 37 };
  const screen = project(scope.matrix, start);
  near(project(inverse, screen), start);
  const moved = project(inverse, { x: screen.x + 17, y: screen.y - 9 });
  near(project(scope.matrix, moved), { x: screen.x + 17, y: screen.y - 9 });
  const roots = shapeScope(loaded, null);
  assert.deepEqual(roots.shapes.map(getShapeId), [getShapeId(outer)]);
  assert.deepEqual(roots.matrix, IDENTITY);
  assert.deepEqual(
    shapeScope(loaded, getShapeId(inner)).shapes.map(getShapeId),
    getGroupChildren(roots.shapes[0]).map(getShapeId),
  );
});

test('singular transforms are non-editable and matrix composition preserves translation order', () => {
  assert.equal(invert([0, 0, 0, 1, 10, 20]), null);
  near(project(compose([2, 0, 0, 3, 10, 20], [1, 0, 0, 1, 5, 7]), { x: 1, y: 2 }), {
    x: 22,
    y: 47,
  });
});
