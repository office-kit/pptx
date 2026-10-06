import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  addSlideTextBox,
  createPresentation,
  getParagraphEndFormat,
  getShapeParagraphElements,
  getShapeXmlString,
  getSlideShapes,
  getSlides,
  getTableCell,
  getTableCellParagraphs,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphs,
  setShapeTextFormat,
  setTableCellParagraphs,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { SHAPE_ELEMENT } from '../src/api/_internal-symbols.ts';
import { NS, attr, elem, firstChildElement, qname } from '../src/internal/xml/index.ts';

const shapeBox = (slide: ReturnType<typeof addBlankSlide>) =>
  addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    text: 'first\nsecond',
  });

function addUnknownParagraphExtension(shape: ReturnType<typeof shapeBox>): void {
  const txBody = firstChildElement(shape[SHAPE_ELEMENT], qname('p', 'txBody', NS.pml));
  const paragraph = txBody?.children.find(
    (child) => child.kind === 'element' && child.name.localName === 'p',
  );
  if (!paragraph || paragraph.kind !== 'element') throw new Error('paragraph missing');
  const end = firstChildElement(paragraph, qname('a', 'endParaRPr', NS.dml))!;
  end.children.push(
    elem(qname('a', 'extLst', NS.dml), {
      children: [
        elem(qname('a', 'ext', NS.dml), {
          attrs: [attr(qname('', 'uri', ''), 'paragraph-end-unknown')],
        }),
      ],
    }),
  );
}

it('formats a shape paragraph end mark without rebuilding runs or unknown XML', async () => {
  const pres = createPresentation();
  const shape = shapeBox(addBlankSlide(pres));
  setShapeParagraphs(shape, [
    { runs: [{ text: 'first', format: { italic: true } }], endFormat: { bold: true, size: 18 } },
    { runs: [{ text: 'second', format: { underline: true } }], endFormat: { font: 'Arial' } },
  ]);
  addUnknownParagraphExtension(shape);
  const beforeRuns = getShapeParagraphElements(shape, 0);

  setShapeTextFormat(shape, { color: '#123456' }, { paragraphEnd: 0 });
  expect(getParagraphEndFormat(shape, 0)).toEqual({ bold: true, size: 18, color: '#123456' });
  expect(getShapeParagraphElements(shape, 0)).toEqual(beforeRuns);
  expect(getShapeXmlString(shape)).toContain('paragraph-end-unknown');

  setShapeTextFormat(shape, { italic: true }, { paragraphEnd: 0, reset: true });
  expect(getParagraphEndFormat(shape, 0)).toEqual({ italic: true });
  expect(getShapeParagraphElements(shape, 0)).toEqual(beforeRuns);

  const reloaded = await loadPresentation(await savePresentation(pres));
  const restored = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
  expect(getParagraphEndFormat(restored, 0)).toEqual({ italic: true });
  expect(getShapeParagraphElements(restored, 0)).toEqual(beforeRuns);
  expect(getShapeXmlString(restored)).toContain('paragraph-end-unknown');
});

it('formats a table-cell paragraph end mark and preserves the other paragraphs', async () => {
  const pres = createPresentation();
  const table = addSlideTable(addBlankSlide(pres), {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(2),
    rows: [['first', 'other']],
  });
  const cell = getTableCell(table, 0, 0)!;
  setTableCellParagraphs(cell, [
    { runs: [{ text: 'first', format: { bold: true } }], endFormat: { size: 15 } },
    { runs: [{ text: 'second', format: { italic: true } }], endFormat: { font: 'Arial' } },
  ]);
  const before = getTableCellParagraphs(cell);

  setTableCellTextFormat(cell, { color: '#ABCDEF' }, { paragraphEnd: 0 });
  expect(getTableCellParagraphs(cell)[0]!.endFormat).toEqual({ size: 15, color: '#ABCDEF' });
  expect(getTableCellParagraphs(cell)[0]!.elements).toEqual(before[0]!.elements);

  setTableCellTextFormat(cell, { underline: true }, { paragraphEnd: 0, reset: true });
  expect(getTableCellParagraphs(cell)[0]!.endFormat).toEqual({ underline: true });
  expect(getTableCellParagraphs(cell)[1]).toEqual(before[1]);

  const reloaded = await loadPresentation(await savePresentation(pres));
  const restoredCell = getTableCell(getSlideShapes(getSlides(reloaded)[0]!)[0]!, 0, 0)!;
  expect(getTableCellParagraphs(restoredCell)[0]!.endFormat).toEqual({ underline: true });
  expect(getTableCellParagraphs(restoredCell)[1]).toEqual(before[1]);
});

it('rejects invalid paragraph-end indices and formats without mutating XML', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = shapeBox(slide);
  const table = addSlideTable(slide, {
    x: inches(1),
    y: inches(4),
    w: inches(4),
    h: inches(1),
    rows: [['cell']],
  });
  const cell = getTableCell(table, 0, 0)!;
  const shapeBefore = getShapeXmlString(shape);
  const cellBefore = getShapeXmlString(table);

  expect(() => setShapeTextFormat(shape, { bold: true }, { paragraphEnd: -1 })).toThrow(RangeError);
  expect(() => setShapeTextFormat(shape, { bold: true }, { paragraphEnd: 99 })).toThrow(RangeError);
  expect(() => setShapeTextFormat(shape, { size: -1 }, { paragraphEnd: 0 })).toThrow();
  expect(getShapeXmlString(shape)).toBe(shapeBefore);

  expect(() => setTableCellTextFormat(cell, { bold: true }, { paragraphEnd: -1 })).toThrow(
    RangeError,
  );
  expect(() => setTableCellTextFormat(cell, { bold: true }, { paragraphEnd: 99 })).toThrow(
    RangeError,
  );
  expect(() => setTableCellTextFormat(cell, { size: -1 }, { paragraphEnd: 0 })).toThrow();
  expect(getShapeXmlString(table)).toBe(cellBefore);
});
