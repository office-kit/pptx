import { expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getTableCells,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setTableCellTextDirection,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

it.each(['vert', 'vert270'] as const)(
  'renders saved %s table text in pure SVG',
  async (direction) => {
    const source = createPresentation();
    const table = addSlideTable(addBlankSlide(source), {
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(3),
      rows: [['Vertical label']],
    });
    setTableCellTextDirection(getTableCells(table)[0]![0]!, direction);
    const pres = await loadPresentation(await savePresentation(source));
    const svg = renderSlideToSvg(pres, getSlides(pres)[0]!, { textLayout: 'svg' });
    expect(svg).toContain(`rotate(${direction === 'vert' ? 90 : 270} `);
    expect(svg).toContain('Vertical label');
    expect(svg).not.toContain('foreignObject');
  },
);
