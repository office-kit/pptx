import { expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

it('creates an aligned paragraph when a covered cell has no text body', async () => {
  const pres = api.createPresentation();
  const slide = api.addBlankSlide(pres);
  const table = api.addSlideTable(slide, {
    x: api.inches(1),
    y: api.inches(1),
    w: api.inches(4),
    h: api.inches(1),
    rows: [['a', 'b']],
  });
  api.mergeTableCells(table, { row: 0, col: 0, rowSpan: 1, colSpan: 2 }, { coveredText: 'drop' });
  api.setTableCellAlignment(api.getTableCell(table, 0, 1), 'center');
  const reloaded = await api.loadPresentation(await api.savePresentation(pres));
  const xml = api.getSlideXmlString(api.getSlides(reloaded)[0]!);
  expect(xml).toContain('<a:pPr algn="ctr"');
  if (isSchemaValidationAvailable()) expectSchemaValid(xml, 'pml');
});
