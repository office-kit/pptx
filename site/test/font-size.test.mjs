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
  addSlideTable,
  getTableCells,
  getTableCellParagraphs,
  setTableCellTextFormat,
  getParagraphEndFormat,
} from '@office-kit/pptx';
import {
  stepFontSize,
  stepShapeFontSize,
  stepTableCellFontSize,
} from '../src/lib/editor/core/font-size.ts';

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

test('font size steps follow Mac PowerPoint including custom values and limits', () => {
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
