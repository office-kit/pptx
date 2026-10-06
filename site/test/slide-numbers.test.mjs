// Deck-wide slide numbers: which shape ends up carrying the number, and what
// the switch is allowed to delete when it goes off again.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  addBlankSlide,
  addSlide,
  addSlideImage,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  findSlidePlaceholder,
  getShapeBounds,
  getShapeBoundsResolved,
  getShapeId,
  getShapeParagraphElements,
  getShapeText,
  getSlideLayoutName,
  getSlideLayouts,
  getSlideShapes,
  getSlideSize,
  inches,
  loadPresentation,
  removeShape,
  setShapeText,
} from '@office-kit/pptx';
import {
  hideSlideNumber,
  showSlideNumber,
  slideNumberShape,
} from '../src/lib/editor/core/slide-numbers.ts';

const fixturePath = fileURLToPath(
  new URL('../../test/fixtures/minimal/blank.pptx', import.meta.url),
);

// `createPresentation`'s layouts reserve no footer slots, so this is the deck
// with nowhere for a number to go but the corner.
const plainDeck = () => {
  const pres = createPresentation();
  return { pres, slide: addBlankSlide(pres) };
};

// A PowerPoint template, whose layouts reserve a `sldNum` slot.
const templateDeck = async () => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = getSlideLayouts(pres).find((l) => getSlideLayoutName(l) === 'Title and Content');
  return { pres, slide: addSlide(pres, { layout }) };
};

const fieldOf = (shape) => getShapeParagraphElements(shape, 0)[0];

// A 1×1 transparent PNG — the smallest thing `addSlideImage` accepts.
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

test('puts a live slide-number field in the corner when nothing reserves a slot', () => {
  const { pres, slide } = plainDeck();

  const shape = showSlideNumber(pres, slide);

  assert.ok(shape);
  assert.equal(fieldOf(shape).kind, 'fld');
  assert.equal(fieldOf(shape).type, 'slidenum');
  const size = getSlideSize(pres);
  const bounds = getShapeBounds(shape);
  assert.ok(bounds.x > size.width / 2 && bounds.y > size.height / 2);
  assert.ok(bounds.x + bounds.w <= size.width && bounds.y + bounds.h <= size.height);
  assert.equal(fieldOf(shape).format?.size, 12);
});

test('fills the template’s own slide-number placeholder instead of adding a box', async () => {
  const { pres, slide } = await templateDeck();
  const placeholder = findSlidePlaceholder(slide, 'sldNum');
  const before = getSlideShapes(slide).length;

  const shape = showSlideNumber(pres, slide);

  assert.equal(getShapeId(shape), getShapeId(placeholder));
  assert.equal(getSlideShapes(slide).length, before);
  assert.equal(fieldOf(shape).type, 'slidenum');
});

test('restores the template’s slot when the author deleted it', async () => {
  const { pres, slide } = await templateDeck();
  const before = getSlideShapes(slide).length;
  removeShape(findSlidePlaceholder(slide, 'sldNum'));

  const shape = showSlideNumber(pres, slide);

  assert.ok(shape);
  // The slot's position comes from the layout and master, not from a box this
  // module guessed at, so the number sits where the template designed it.
  assert.ok(getShapeBoundsResolved(pres, shape));
  assert.equal(getShapeBounds(shape), null);
  assert.equal(getSlideShapes(slide).length, before);
});

test('is idempotent — a second call does not add a second number', () => {
  const { pres, slide } = plainDeck();

  const first = showSlideNumber(pres, slide);
  const again = showSlideNumber(pres, slide);

  assert.equal(getShapeId(again), getShapeId(first));
  assert.equal(getSlideShapes(slide).length, 1);
});

// A picture or a table has no text body of its own, and asking one for its
// paragraphs throws. The switch reads every shape on the slide, so it has to
// step past them — the editor renders this state on every keystroke.
test('looks past pictures and tables, which have no text body of their own', () => {
  const { pres, slide } = plainDeck();
  addSlideImage(slide, PNG, { x: inches(0), y: inches(0), w: inches(1), h: inches(1) });
  addSlideTable(slide, {
    x: inches(2),
    y: inches(2),
    w: inches(3),
    h: inches(1),
    rows: [
      ['A', 'B'],
      ['C', 'D'],
    ],
  });

  assert.equal(slideNumberShape(slide), null);
  assert.ok(showSlideNumber(pres, slide));
  assert.equal(hideSlideNumber(slide), true);
  assert.equal(getSlideShapes(slide).length, 2);
});

test('off removes the number and nothing else', () => {
  const { pres, slide } = plainDeck();
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
    text: 'Body',
  });
  showSlideNumber(pres, slide);

  assert.equal(hideSlideNumber(slide), true);

  assert.equal(slideNumberShape(slide), null);
  assert.deepEqual(getSlideShapes(slide).map(getShapeText), ['Body']);
});

test('leaves a box that merely mentions a number among its text alone', () => {
  const { slide } = plainDeck();
  const box = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(2),
    h: inches(1),
    text: '',
  });
  setShapeText(box, 'Page 3 of 9');

  assert.equal(slideNumberShape(slide), null);
  assert.equal(hideSlideNumber(slide), false);
  assert.equal(getSlideShapes(slide).length, 1);
});

test('reports nothing to remove on a slide that never had a number', () => {
  const { slide } = plainDeck();
  assert.equal(hideSlideNumber(slide), false);
});
