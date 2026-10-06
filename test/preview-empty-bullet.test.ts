import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  getTableCells,
  setShapeParagraphs,
  setTableCellParagraphs,
  setParagraphBullet,
  inches,
  loadPresentation,
  savePresentation,
  getSlides,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

for (const kind of ['shape', 'table'] as const)
  for (const textLayout of ['svg', 'foreignObject'] as const)
    it(`${kind}: hides empty paragraph bullets without losing populated bullets (${textLayout})`, async () => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const bounds = { x: inches(1), y: inches(1), w: inches(6), h: inches(3) };
      const paragraphs = [{ runs: [{ text: 'Before' }] }, { runs: [] }];
      if (kind === 'shape') {
        const shape = addSlideTextBox(slide, { ...bounds, text: '' });
        setShapeParagraphs(shape, paragraphs);
        setParagraphBullet(shape, 0, { char: '◆' });
        setParagraphBullet(shape, 1, { char: '◆' });
      } else {
        const table = addSlideTable(slide, { ...bounds, rows: [['']] });
        const cell = getTableCells(table)[0]![0]!;
        setTableCellParagraphs(cell, paragraphs);
        setParagraphBullet(cell, 0, { char: '◆' });
        setParagraphBullet(cell, 1, { char: '◆' });
      }
      for (const deck of [pres, await loadPresentation(await savePresentation(pres))]) {
        const rendered = renderSlideToSvg(deck, getSlides(deck)[0]!, { textLayout });
        expect(rendered.match(/◆/g)).toHaveLength(1);
        expect(rendered).toContain('Before');
      }
    });
