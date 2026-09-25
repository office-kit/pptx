import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  getShapeRunFormatEffective,
  getShapeXmlString,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  getTableCells,
  getSlides,
  getSlideShapes,
  inches,
  loadPresentation,
  savePresentation,
  setShapeParagraphs,
  setTableCellParagraphs,
} from '../src/api/index.ts';
import { textFormatsInRange } from '../site/src/lib/editor/core/text-format-selection.ts';
import { toggleTextFormat } from '../site/src/lib/editor/core/text-format-toggle.ts';
import { projectTextEdits } from '../site/src/lib/editor/core/text-edit-preview.ts';

for (const kind of ['shape', 'cell'] as const)
  describe(`${kind} caret formats`, () => {
    it('reads empty paragraph end formats before and after save/reload and pending input', async () => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
      const shape =
        kind === 'shape'
          ? addSlideTextBox(slide, { ...bounds, text: '' })
          : addSlideTable(slide, { ...bounds, rows: [['', 'Other']] });
      const position = kind === 'cell' ? { row: 0, col: 0 } : undefined;
      const paragraphs = [
        { runs: [], endFormat: { bold: true, size: 32 } },
        { runs: [{ text: '日本語 English', format: { italic: true } }] },
        { runs: [], endFormat: { bold: false, size: 18 } },
      ];
      if (kind === 'shape') setShapeParagraphs(shape, paragraphs);
      else setTableCellParagraphs(getTableCells(shape)[0]![0]!, paragraphs);
      const loaded = await loadPresentation(await savePresentation(pres));
      for (const target of [shape, getSlideShapes(getSlides(loaded)[0]!)[0]!]) {
        const formats = textFormatsInRange(target, { start: 0, end: 0 }, position);
        expect(formats).toEqual([{ bold: true, size: 32 }]);
        expect(toggleTextFormat(formats, 'bold')).toEqual({ bold: false });
        expect(textFormatsInRange(target, { start: 13, end: 13 }, position)).toEqual([
          { bold: false, size: 18 },
        ]);
        expect(textFormatsInRange(target, { start: 1, end: 1 }, position)).toEqual([
          { italic: true },
        ]);
        expect(textFormatsInRange(target, { start: 1, end: 12 }, position)).toEqual([
          { italic: true },
        ]);
        const pending = projectTextEdits(target, [{ start: 1, end: 1, text: '追加' }], position);
        expect(textFormatsInRange(pending, { start: 15, end: 15 }, position)).toEqual([
          { bold: false, size: 18 },
        ]);
      }
    });
  });

it('uses inherited placeholder styles for toolbar selections after pending edits', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
  const shape = findSlidePlaceholder(slide, 'body')!;
  setShapeParagraphs(shape, [{ runs: [{ text: 'English' }], endFormat: { bold: true } }]);
  const xml = getShapeXmlString(shape);
  const inherited = getShapeRunFormatEffective(pres, shape, 0, 0);
  expect(inherited.size).toBeGreaterThan(18);
  expect(inherited.bold).toBe(true);
  for (const target of [
    shape,
    projectTextEdits(shape, [{ start: 0, end: 0, text: '日本語\n' }], undefined, pres),
  ]) {
    const context = { pres, source: shape };
    const caret = textFormatsInRange(target, { start: 5, end: 5 }, undefined, context);
    expect(caret[0]).toMatchObject(inherited);
    expect(toggleTextFormat(caret, 'bold')).toEqual({ bold: false });
    expect(
      textFormatsInRange(target, { start: 4, end: 7 }, undefined, context).every(
        (format) => format.size === inherited.size,
      ),
    ).toBe(true);
  }
  expect(getShapeXmlString(shape)).toBe(xml);
});
