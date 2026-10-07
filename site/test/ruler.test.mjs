import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getParagraphIndent,
  getParagraphPropertiesEffective,
  setParagraphIndent,
  setParagraphTabs,
  setShapeParagraphs,
} from '@office-kit/pptx';
import { applyRulerChange, indentAfterDrag, rulerAxis } from '../src/lib/editor/core/ruler.ts';
import { editTabStops } from '../src/lib/editor/core/paragraph-tabs.ts';

const close = (actual, expected) => {
  for (const key of Object.keys(expected)) {
    if (typeof expected[key] === 'object') close(actual[key], expected[key]);
    else if (typeof expected[key] === 'string') assert.equal(actual[key], expected[key]);
    else
      assert.ok(
        Math.abs(actual[key] - expected[key]) < 1e-9,
        `${key}: ${actual[key]} != ${expected[key]}`,
      );
  }
};

test('unrotated horizontal text measures from its own origin', () => {
  close(
    rulerAxis({
      origin: { x: 110, y: 40 },
      direction: { x: 2, y: 0 },
      center: { x: 200, y: 100 },
      flow: 'horizontal',
    }),
    { axis: 'x', origin: { x: 110, y: 40 }, scale: 2 },
  );
});

test('rotated text is measured as if its shape were unrotated about the center', () => {
  // A box centred at (200,100), 180px wide, origin 90px left of centre and
  // 60px above, rotated 90° clockwise: the origin lands right of centre.
  const angle = Math.PI / 2;
  const rotate = (x, y) => ({
    x: x * Math.cos(angle) - y * Math.sin(angle),
    y: x * Math.sin(angle) + y * Math.cos(angle),
  });
  const offset = rotate(-90, -60);
  const direction = rotate(1.5, 0);
  close(
    rulerAxis({
      origin: { x: 200 + offset.x, y: 100 + offset.y },
      direction,
      center: { x: 200, y: 100 },
      flow: 'horizontal',
    }),
    { axis: 'x', origin: { x: 110, y: 40 }, scale: 1.5 },
  );
  // Any angle, including reflection-compensating half turns.
  for (const degrees of [30, 135, 180, -45, 270]) {
    const a = (degrees * Math.PI) / 180;
    const r = (x, y) => ({
      x: x * Math.cos(a) - y * Math.sin(a),
      y: x * Math.sin(a) + y * Math.cos(a),
    });
    const o = r(-90, -60);
    close(
      rulerAxis({
        origin: { x: 200 + o.x, y: 100 + o.y },
        direction: r(1, 0),
        center: { x: 200, y: 100 },
        flow: 'horizontal',
      }),
      { axis: 'x', origin: { x: 110, y: 40 }, scale: 1 },
    );
  }
});

test('vertical text measures down the vertical ruler, and upward for vert270', () => {
  close(
    rulerAxis({
      origin: { x: 280, y: 20 },
      direction: { x: 0, y: 1 },
      center: { x: 200, y: 100 },
      flow: 'vertical',
    }),
    { axis: 'y', origin: { x: 280, y: 20 }, scale: 1 },
  );
  close(
    rulerAxis({
      origin: { x: 120, y: 180 },
      direction: { x: 0, y: -1 },
      center: { x: 200, y: 100 },
      flow: 'vertical-reversed',
    }),
    { axis: 'y', origin: { x: 120, y: 180 }, scale: -1 },
  );
  // vert text in a shape rotated 90° runs right-to-left on screen.
  close(
    rulerAxis({
      origin: { x: 280, y: 180 },
      direction: { x: -2, y: 0 },
      center: { x: 200, y: 100 },
      flow: 'vertical',
    }),
    { axis: 'y', origin: { x: 280, y: 20 }, scale: 2 },
  );
});

test('indent handles keep PowerPoint margin rules', () => {
  assert.deepEqual(indentAfterDrag({ left: 0, first: 0 }, 'first', 358775), {
    firstLineEmu: 358775,
  });
  // Hanging from zero: marL moves, the first line stays where it was.
  assert.deepEqual(indentAfterDrag({ left: 0, first: 0 }, 'hanging', 358775), {
    leftEmu: 358775,
    firstLineEmu: -358775,
  });
  assert.deepEqual(indentAfterDrag({ left: 100, first: -50 }, 'left', -500), { leftEmu: 0 });
  assert.deepEqual(indentAfterDrag({ left: 0, first: 0 }, 'first', -1e9), {
    firstLineEmu: -51206400,
  });
});

test('mixed paragraphs move relative to their own indents and tabs', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 3000000,
    h: 1000000,
    text: '',
  });
  setShapeParagraphs(shape, [{ runs: [{ text: 'a' }] }, { runs: [{ text: 'b' }] }]);
  setParagraphIndent(shape, 0, { leftEmu: 100000, firstLineEmu: 0 });
  setParagraphIndent(shape, 1, { leftEmu: 300000, firstLineEmu: 50000 });
  setParagraphTabs(shape, 0, { tabStops: [{ positionEmu: 360000, alignment: 'left' }] });
  setParagraphTabs(shape, 1, { tabStops: [{ positionEmu: 720000, alignment: 'right' }] });
  applyRulerChange(pres, shape, [0, 1], { kind: 'indent', handle: 'left', delta: 20000 });
  assert.equal(getParagraphIndent(shape, 0).leftEmu, 120000);
  assert.equal(getParagraphIndent(shape, 1).leftEmu, 320000);
  assert.equal(getParagraphIndent(shape, 1).firstLineEmu, 50000);
  // Moving the first paragraph's stop leaves the second paragraph's alone.
  applyRulerChange(pres, shape, [0, 1], {
    kind: 'tabs',
    edits: [{ kind: 'move', fromEmu: 360000, toEmu: 400000 }],
  });
  assert.deepEqual(getParagraphPropertiesEffective(pres, shape, 0).tabStops, [
    { positionEmu: 400000, alignment: 'left' },
  ]);
  assert.deepEqual(getParagraphPropertiesEffective(pres, shape, 1).tabStops, [
    { positionEmu: 720000, alignment: 'right' },
  ]);
});

test('moving a stop keeps each paragraph’s own alignment', () => {
  assert.deepEqual(
    editTabStops(
      [
        { positionEmu: 10, alignment: 'decimal' },
        { positionEmu: 30, alignment: 'left' },
      ],
      [{ kind: 'move', fromEmu: 10, toEmu: 40 }],
    ),
    [
      { positionEmu: 30, alignment: 'left' },
      { positionEmu: 40, alignment: 'decimal' },
    ],
  );
  assert.deepEqual(
    editTabStops(
      [{ positionEmu: 30, alignment: 'left' }],
      [{ kind: 'move', fromEmu: 10, toEmu: 40 }],
    ),
    [{ positionEmu: 30, alignment: 'left' }],
  );
});
