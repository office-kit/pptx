import { test } from 'node:test';
import assert from 'node:assert/strict';
import { snapMove, smartGuideTargets } from '../src/smart-guides.ts';
const shape = (id, x, y, w, h, rotation = 0) => ({
  id,
  bounds: { x, y, w, h },
  rotation,
  connections: {},
});
const slide = { width: 1000, height: 800 };
test('moved and deleted drawing guides determine alignment targets', () => {
  const moving = shape(1, 30, 40, 100, 60);
  const options = { smart: false, drawing: true, guides: [{ axis: 'x', offset: 120 }] };
  const result = snapMove([moving], [], slide, { x: 587, y: 327 }, 5, undefined, options);
  assert.equal(result.x, 590);
  assert.equal(result.y, 327);
  assert.deepEqual(result.guides, [{ axis: 'x', value: 620, from: 0, to: 800 }]);
  options.guides = [];
  assert.deepEqual(snapMove([moving], [], slide, { x: 587, y: 327 }, 5, undefined, options), {
    x: 587,
    y: 327,
    guides: [],
  });
});
test('guide settings disable smart alignment while drawing guides independently snap to center', () => {
  const moving = shape(1, 30, 40, 100, 60);
  const target = shape(2, 300, 300, 100, 60);
  const settings = { smart: false, drawing: false };
  assert.deepEqual(
    snapMove([moving], [target], slide, { x: 267, y: 262 }, 5, undefined, settings),
    {
      x: 267,
      y: 262,
      guides: [],
    },
  );
  settings.drawing = true;
  const centered = snapMove([moving], [target], slide, { x: 417, y: 327 }, 5, undefined, settings);
  assert.equal(centered.x, 420);
  assert.equal(centered.y, 330);
  const nearSlideEdge = snapMove([moving], [], slide, { x: -28, y: -38 }, 5, undefined, settings);
  assert.deepEqual(nearSlideEdge, { x: -28, y: -38, guides: [] });
});
test('smart guides choose nearest object alignment independently on each axis', () => {
  const moving = shape(1, 30, 40, 100, 60);
  const target = shape(2, 300, 300, 100, 60);
  const result = snapMove([moving], [moving, target], slide, { x: 267, y: 262 }, 5);
  assert.equal(result.x, 270);
  assert.equal(result.y, 260);
  assert.deepEqual(result.guides, [
    { axis: 'x', value: 300, from: 300, to: 360 },
    { axis: 'y', value: 300, from: 300, to: 400 },
  ]);
  assert.deepEqual(
    result.guides.map((g) => [g.axis, g.value]),
    [
      ['x', 300],
      ['y', 300],
    ],
  );
  const free = snapMove([moving], [moving, target], slide, { x: 250, y: 240 }, 5);
  assert.deepEqual(free, { x: 250, y: 240, guides: [] });
});
test('multi-selection snaps as one box without changing relative distances', () => {
  const first = shape(1, 100, 100, 100, 50);
  const second = shape(2, 300, 200, 100, 50);
  const result = snapMove([first, second], [first, second], slide, { x: 247, y: 50 }, 5);
  assert.equal(result.x, 250);
  assert.equal(result.guides[0].value, 500);
  assert.equal(second.bounds.x + result.x - (first.bounds.x + result.x), 200);
});
test('rotated visible bounds align and Shift preserves its constrained axis', () => {
  const moving = shape(1, 100, 100, 200, 100, 90);
  const target = shape(2, 400, 400, 100, 100);
  const result = snapMove([moving], [target], slide, { x: 247, y: 347 }, 5, 'x');
  assert.equal(result.x, 250);
  assert.equal(result.y, 347);
  assert.equal(result.guides.length, 1);
  assert.equal(result.guides[0].axis, 'x');
});

test('equal spacing uses edge gaps for different widths and extends an existing row', () => {
  const moving = shape(1, 0, 210, 40, 50);
  const others = [shape(2, 100, 200, 60, 70), shape(3, 420, 200, 100, 70)];
  const between = snapMove([moving], others, slide, { x: 267, y: 0 }, 5, 'x');
  assert.equal(between.x, 270);
  assert.deepEqual(between.guides, [
    { kind: 'spacing', axis: 'x', value: 235, from: 160, to: 270 },
    { kind: 'spacing', axis: 'x', value: 235, from: 310, to: 420 },
  ]);
  const after = snapMove([moving], others, slide, { x: 777, y: 0 }, 5, 'x');
  assert.equal(after.x, 780);
  assert.deepEqual(
    after.guides.map((g) => g.to - g.from),
    [260, 260],
  );
});

test('vertical spacing moves a selection together and respects axis constraints', () => {
  const moving = [shape(1, 210, 0, 40, 20), shape(2, 210, 30, 40, 10)];
  const others = [shape(3, 200, 100, 70, 60), shape(4, 200, 420, 70, 100)];
  const result = snapMove(moving, others, slide, { x: 0, y: 267 }, 5, 'y');
  assert.equal(result.y, 270);
  assert.deepEqual(
    result.guides.map((g) => [g.kind, g.axis, g.to - g.from]),
    [
      ['spacing', 'y', 110],
      ['spacing', 'y', 110],
    ],
  );
  assert.equal(snapMove(moving, others, slide, { x: 0, y: 267 }, 5, 'x').y, 267);
});

test('spacing ignores unrelated rows, obstructed gaps, and disabled smart guides', () => {
  const moving = shape(1, 0, 210, 40, 50);
  const others = [shape(2, 100, 200, 60, 70), shape(3, 420, 200, 100, 70)];
  const distant = others.map((item) => ({ ...item, bounds: { ...item.bounds, y: 500 } }));
  assert.equal(snapMove([moving], distant, slide, { x: 267, y: 0 }, 5, 'x').x, 267);
  const blocked = [...others, shape(4, 220, 200, 10, 70)];
  assert.equal(snapMove([moving], blocked, slide, { x: 267, y: 0 }, 5, 'x').x, 267);
  assert.equal(
    snapMove([moving], others, slide, { x: 267, y: 0 }, 5, 'x', { smart: false, drawing: false }).x,
    267,
  );
  assert.equal(snapMove([moving], others, slide, { x: 264, y: 0 }, 5, 'x').x, 264);
});

test('group child snapping uses slide coordinates for rotated and scaled parents', () => {
  const moving = { ...shape(1, 20, 40, 30, 50), parentTransform: [0, 2, -3, 0, 700, 100] };
  // The projected child occupies x=430..580, y=140..200.
  const target = shape(2, 700, 300, 80, 60);
  const result = snapMove([moving], [target], { width: 2000, height: 1500 }, { x: 117, y: 157 }, 5);
  assert.equal(result.x, 120);
  assert.equal(result.y, 160);
  assert.ok(result.guides.some((g) => g.axis === 'x' && g.value === 700));
  assert.ok(result.guides.some((g) => g.axis === 'y' && g.value === 300));
  const constrained = snapMove([moving], [target], slide, { x: 117, y: 157 }, 5, 'x');
  assert.equal(constrained.y, 157);
});
test('snap candidates exclude selected subtrees and ancestors but include siblings', () => {
  const child = shape(3, 0, 0, 10, 10),
    sibling = shape(4, 20, 20, 10, 10);
  const nested = { ...shape(2, 0, 0, 30, 30), children: [child, sibling] };
  const outerSibling = shape(5, 50, 50, 10, 10);
  const group = { ...shape(1, 0, 0, 100, 100), children: [nested, outerSibling] };
  const external = { ...shape(6, 200, 200, 50, 50), children: [shape(7, 0, 0, 10, 10)] };
  assert.deepEqual(
    smartGuideTargets([group, external], [child]).map((s) => s.id),
    [4, 5, 6],
  );
  assert.deepEqual(
    smartGuideTargets([group, external], [nested]).map((s) => s.id),
    [5, 6],
  );
  assert.deepEqual(
    smartGuideTargets([group, external], [group]).map((s) => s.id),
    [6],
  );
  assert.deepEqual(
    smartGuideTargets([group, external], [child, sibling]).map((s) => s.id),
    [5, 6],
  );
});

test('grid snapping preserves selection offsets, constrained axes and smart alignment priority', () => {
  const moving = [shape(1, 31, 43, 100, 60), shape(2, 140, 85, 50, 30)];
  const options = { smart: false, drawing: false, grid: { x: 20, y: 25 } };
  assert.deepEqual(snapMove(moving, [], slide, { x: 14, y: 10 }, 5, undefined, options), {
    x: 9,
    y: 7,
    guides: [],
  });
  assert.deepEqual(snapMove(moving, [], slide, { x: 14, y: 0 }, 5, 'x', options), {
    x: 9,
    y: 0,
    guides: [],
  });
  const aligned = snapMove(
    [moving[0]],
    [shape(3, 46, 300, 100, 60)],
    slide,
    { x: 14, y: 10 },
    5,
    undefined,
    { ...options, smart: true },
  );
  assert.equal(aligned.x, 15);
  assert.equal(aligned.y, 7);
});
