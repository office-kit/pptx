import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  createPresentation,
  inches,
  setShapeTextFormat,
  setTableCellTextFormat,
  getTableCells,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

describe('text highlight preview', () => {
  it.each(['svg', 'foreignObject'] as const)(
    'renders and clears shape and table highlights in %s mode',
    (textLayout) => {
      const pres = createPresentation();
      const slide = addBlankSlide(pres);
      const shape = addSlideTextBox(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(1),
        text: 'English 日本語',
      });
      const table = addSlideTable(slide, {
        x: inches(1),
        y: inches(3),
        w: inches(4),
        h: inches(1),
        rows: [['Cell', 'Other']],
      });
      const cell = getTableCells(table)[0]![0]!;
      setShapeTextFormat(shape, { highlight: '#FFCC00' }, { range: { start: 0, end: 7 } });
      setTableCellTextFormat(cell, { highlight: '#00CCFF' });
      const svg = renderSlideToSvg(pres, slide, { textLayout });
      for (const color of ['#FFCC00', '#00CCFF']) {
        expect(svg).toContain(
          textLayout === 'svg' ? `fill="${color}"` : `background-color:${color}`,
        );
      }
      setShapeTextFormat(shape, { highlight: null });
      setTableCellTextFormat(cell, { highlight: null });
      const cleared = renderSlideToSvg(pres, slide, { textLayout });
      expect(cleared).not.toContain('#FFCC00');
      expect(cleared).not.toContain('#00CCFF');
    },
  );
});
