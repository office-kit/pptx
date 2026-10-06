import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  getSlides,
  getSlideShapes,
  getTableCells,
  getSlideXmlString,
  getParagraphPropertiesEffective,
  setParagraphTypography,
  setParagraphAlignment,
  savePresentation,
  loadPresentation,
  inches,
} from '../src/api/index.ts';

it('round-trips typography per paragraph in shapes and cells, preserving unrelated properties', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
  const shape = addSlideTextBox(slide, { ...bounds, text: 'First\nSecond' });
  const table = addSlideTable(slide, { ...bounds, rows: [['First\nSecond', 'Other']] });
  const cell = getTableCells(table)[0]![0]!;
  const typography = {
    asianLineBreak: false,
    latinLineBreak: true,
    hangingPunctuation: false,
    fontAlignment: 'baseline' as const,
  };
  for (const target of [shape, cell]) {
    setParagraphAlignment(target, 1, 'right');
    setParagraphTypography(target, 1, typography);
    expect(getParagraphPropertiesEffective(pres, target, 1)).toMatchObject({
      ...typography,
      align: 'right',
    });
    expect(getParagraphPropertiesEffective(pres, target, 0).latinLineBreak).toBeUndefined();
  }
  const xml = getSlideXmlString(slide);
  expect(xml).toContain('eaLnBrk="0"');
  expect(xml).toContain('latinLnBrk="1"');
  expect(xml).toContain('hangingPunct="0"');
  expect(xml).toContain('fontAlgn="base"');
  const loaded = await loadPresentation(await savePresentation(pres));
  const [loadedShape, loadedTable] = getSlideShapes(getSlides(loaded)[0]!);
  for (const target of [loadedShape!, getTableCells(loadedTable!)[0]![0]!]) {
    expect(getParagraphPropertiesEffective(loaded, target, 1)).toMatchObject(typography);
    setParagraphTypography(target, 1, { latinLineBreak: false });
    expect(getParagraphPropertiesEffective(loaded, target, 1)).toMatchObject({
      ...typography,
      latinLineBreak: false,
    });
    setParagraphTypography(target, 1, { fontAlignment: null, asianLineBreak: null });
    const cleared = getParagraphPropertiesEffective(loaded, target, 1);
    expect(cleared.fontAlignment).toBeUndefined();
    expect(cleared.asianLineBreak).toBeUndefined();
    expect(cleared.align).toBe('right');
    expect(cleared.hangingPunctuation).toBe(false);
  }
});
