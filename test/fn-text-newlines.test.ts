import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  inches,
  getTableCells,
  setShapeText,
  setTableCellText,
  getShapeParagraphCount,
  getTableCellParagraphs,
  getShapeParagraphElements,
  getSlides,
  getSlideShapes,
  loadPresentation,
  savePresentation,
} from '../src/api/index.ts';

for (const kind of ['shape', 'table'] as const)
  for (const mode of ['replace', 'range', 'preserve'] as const)
    it(`${kind}: ${mode} inserts inline breaks and round-trips`, async () => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const bounds = { x: inches(1), y: inches(1), w: inches(5), h: inches(3) };
      const shape =
        kind === 'shape'
          ? addSlideTextBox(slide, { ...bounds, text: 'AB' })
          : addSlideTable(slide, { ...bounds, rows: [['AB']] });
      const options = {
        newlines: 'break' as const,
        ...(mode === 'range' ? { range: { start: 1, end: 1 } } : {}),
        ...(mode === 'preserve' ? { preserveFormatting: true } : {}),
      };
      const value = mode === 'range' ? '\n\n' : 'A\n\nB';
      if (kind === 'shape') setShapeText(shape, value, options);
      else setTableCellText(getTableCells(shape)[0]![0]!, value, options);
      for (const deck of [pres, await loadPresentation(await savePresentation(pres))]) {
        const result = getSlideShapes(getSlides(deck)[0]!)[0]!;
        const paragraphs =
          kind === 'shape'
            ? getShapeParagraphCount(result)
            : getTableCellParagraphs(getTableCells(result)[0]![0]!).length;
        expect(paragraphs).toBe(1);
        const elements =
          kind === 'shape'
            ? getShapeParagraphElements(result, 0)
            : getTableCellParagraphs(getTableCells(result)[0]![0]!)[0]!.elements;
        expect(elements.map((e) => e.kind)).toEqual(['r', 'br', 'br', 'r']);
      }
    });
