import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBlankSlide,
  addSlideTextBox,
  createPresentation,
  getShapeParagraphElements,
  getShapeText,
  getSlides,
  getSlideShapes,
  loadPresentation,
  savePresentation,
  setShapeTextFormat,
  setShapeParagraphs,
  addSlideTable,
  getTableCells,
  getTableCellParagraphs,
  setTableCellTextFormat,
  setTableCellParagraphs,
  getParagraphEndFormat,
} from '@office-kit/pptx';
import { stepFontSize, stepShapeFontSize, stepTableCellFontSize } from '../src/core/font-size.ts';

import { textFormatsInRange } from '../src/core/text-format-selection.ts';

test('selection sizes match rendering defaults without masking mixed sizes or writing them', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [{ text: 'Default' }, { text: 'Explicit', format: { size: 24 } }] },
  ]);
  const before = getShapeParagraphElements(shape, 0);
  const range = { start: 0, end: getShapeText(shape).length };
  assert.deepEqual(
    textFormatsInRange(shape, range, undefined, { pres }).map((f) => f.size),
    [18, 24],
  );
  assert.deepEqual(
    textFormatsInRange(shape, range).map((f) => f.size),
    [undefined, 24],
  );
  assert.deepEqual(getShapeParagraphElements(shape, 0), before);
});

test('empty paragraphs step their own end formats and respect a selection', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [], endFormat: { size: 20 } },
    { runs: [], endFormat: { size: 44 } },
    { runs: [{ text: 'x', format: { size: 10 } }] },
  ]);
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getParagraphEndFormat(shape, 0).size, 24);
  assert.equal(getParagraphEndFormat(shape, 1).size, 48);
  stepShapeFontSize(pres, shape, 1, shape, { start: 1, end: 2 });
  assert.equal(getParagraphEndFormat(shape, 0).size, 24);
  assert.equal(getParagraphEndFormat(shape, 1).size, 54);
});

test('empty table paragraphs step their end formats', () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    rows: [['']],
  });
  const cell = getTableCells(table)[0][0];
  setTableCellParagraphs(cell, [{ runs: [], endFormat: { size: 20 } }]);
  stepTableCellFontSize(cell, 1);
  assert.equal(getTableCellParagraphs(cell)[0].endFormat.size, 24);
});

test('empty text retains its authored size when stepped repeatedly', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: '',
  });
  setShapeTextFormat(shape, { size: 44 });
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getParagraphEndFormat(shape, 0).size, 48);
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getParagraphEndFormat(shape, 0).size, 54);
});

test('font size steps follow the reference desktop app (Mac) including custom values and limits', () => {
  for (const [size, direction, expected] of [
    [10, 1, 10.5],
    [13, 1, 14],
    [13, -1, 12],
    [44, -1, 40],
    [96, 1, 115],
    [98, 1, 118],
    [100, 1, 120],
    [100, -1, 96],
    [7.5, -1, 7],
    [7, -1, 6],
    [7.1, -1, 7],
    [6, 1, 7],
    [1, -1, 1],
    [4000, 1, 4000],
  ])
    assert.equal(stepFontSize(size, direction), expected, `${size}, ${direction}`);
});

test('table font stepping preserves mixed sizes and untouched characters', () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    rows: [['abcd']],
  });
  const cell = getTableCells(table)[0][0];
  setTableCellTextFormat(cell, { size: 20, bold: true });
  setTableCellTextFormat(cell, { size: 44 }, { range: { start: 2, end: 4 } });
  stepTableCellFontSize(cell, 1, { start: 1, end: 3 });
  assert.deepEqual(
    getTableCellParagraphs(cell)[0].elements.flatMap((element) =>
      [...element.text].map((text) => [text, element.format.size, element.format.bold]),
    ),
    [
      ['a', 20, true],
      ['b', 24, true],
      ['c', 48, true],
      ['d', 44, true],
    ],
  );
});

test('partial font stepping skips unselected runs and paragraphs and survives saving', async () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: 'abCD\nefGH',
  });
  setShapeTextFormat(shape, { size: 20, bold: true });
  setShapeTextFormat(shape, { size: 44, italic: true }, { range: { start: 7, end: 9 } });
  stepShapeFontSize(pres, shape, 1, shape, { start: 6, end: 8 });
  const sizes = (target) =>
    [0, 1].map((index) =>
      getShapeParagraphElements(target, index).flatMap((element) =>
        element.kind === 'r'
          ? [...element.text].map((text) => ({
              text,
              size: element.format.size,
              bold: element.format.bold,
              italic: element.format.italic === true,
            }))
          : [],
      ),
    );
  const expected = [
    [...'abCD'].map((text) => ({ text, size: 20, bold: true, italic: false })),
    [
      { text: 'e', size: 20, bold: true, italic: false },
      { text: 'f', size: 24, bold: true, italic: false },
      { text: 'G', size: 48, bold: true, italic: true },
      { text: 'H', size: 44, bold: true, italic: true },
    ],
  ];
  assert.deepEqual(sizes(shape), expected);
  const loaded = await loadPresentation(await savePresentation(pres));
  const restored = getSlideShapes(getSlides(loaded)[0])[0];
  assert.equal(getShapeText(restored), 'abCD\nefGH');
  assert.deepEqual(sizes(restored), expected);
});

test('zero-length runs retain their authored size when an empty paragraph is stepped', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: '',
  });
  setShapeParagraphs(shape, [{ runs: [{ text: '', format: { size: 44, italic: true } }] }]);
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getParagraphEndFormat(shape, 0).size, 48);
  assert.equal(getShapeParagraphElements(shape, 0)[0].format.italic, true);
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getParagraphEndFormat(shape, 0).size, 54);
});

test('empty runs between visible paragraphs do not shift effective font sizes', () => {
  const pres = createPresentation();
  const shape = addSlideTextBox(addBlankSlide(pres), {
    x: 0,
    y: 0,
    w: 4000000,
    h: 2000000,
    text: '',
  });
  setShapeParagraphs(shape, [
    { runs: [{ text: 'first', format: { size: 20 } }] },
    { runs: [{ text: '', format: { size: 44 } }] },
    { runs: [{ text: 'last', format: { size: 10 } }] },
  ]);
  stepShapeFontSize(pres, shape, 1);
  assert.equal(getShapeParagraphElements(shape, 0)[0].format.size, 24);
  assert.equal(getParagraphEndFormat(shape, 1).size, 48);
  assert.equal(getShapeParagraphElements(shape, 2)[0].format.size, 10.5);
});
