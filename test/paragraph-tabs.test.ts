import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  getTableCells,
  getParagraphPropertiesEffective,
  setParagraphTabs,
  setParagraphSpacing,
  getSlideXmlString,
  savePresentation,
  loadPresentation,
  getSlides,
  getSlideShapes,
  addSlide,
  findSlideLayout,
  findSlidePlaceholder,
  inches,
} from '../src/api/index.ts';

it('round-trips sorted tab stops in text boxes and cells without changing adjacent paragraphs', async () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const bounds = { x: inches(1), y: inches(1), w: inches(4), h: inches(2) };
  const shape = addSlideTextBox(slide, { ...bounds, text: 'One\tTwo\nOther' });
  const table = addSlideTable(slide, { ...bounds, rows: [['One\tTwo\nOther']] });
  const stops = [
    { positionEmu: inches(1), alignment: 'left' as const },
    { positionEmu: inches(2), alignment: 'center' as const },
    { positionEmu: inches(3), alignment: 'right' as const },
    { positionEmu: inches(4), alignment: 'decimal' as const },
  ];
  for (const target of [shape, getTableCells(table)[0]![0]!]) {
    setParagraphSpacing(target, 0, { beforePts: 12 });
    setParagraphTabs(target, 0, { tabStops: [...stops].reverse(), defaultTabSizeEmu: inches(0.5) });
    expect(getParagraphPropertiesEffective(pres, target, 1).tabStops).toBeUndefined();
  }
  const xml = getSlideXmlString(slide);
  expect(xml.indexOf('<a:spcBef>')).toBeLessThan(xml.indexOf('<a:tabLst>'));
  for (const token of ['l', 'ctr', 'r', 'dec']) expect(xml).toContain(`algn="${token}"`);
  const loaded = await loadPresentation(await savePresentation(pres));
  const [loadedShape, loadedTable] = getSlideShapes(getSlides(loaded)[0]!);
  for (const target of [loadedShape!, getTableCells(loadedTable!)[0]![0]!]) {
    expect(getParagraphPropertiesEffective(loaded, target, 0)).toMatchObject({
      tabStops: stops,
      defaultTabSizeEmu: inches(0.5),
      spcBefPts: 12,
    });
    setParagraphTabs(target, 0, { tabStops: [] });
    expect(getParagraphPropertiesEffective(loaded, target, 0)).toMatchObject({
      tabStops: [],
      defaultTabSizeEmu: inches(0.5),
    });
    setParagraphTabs(target, 0, { tabStops: null, defaultTabSizeEmu: null });
    expect(getParagraphPropertiesEffective(loaded, target, 0).tabStops).toBeUndefined();
  }
});

it('restores inherited default tab spacing and rejects invalid updates atomically', async () => {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Title and Content')! });
  const shape = findSlidePlaceholder(slide, 'body')!;
  expect(getParagraphPropertiesEffective(pres, shape, 0).defaultTabSizeEmu).toBe(457200);
  setParagraphTabs(shape, 0, { defaultTabSizeEmu: inches(1) });
  expect(getParagraphPropertiesEffective(pres, shape, 0).defaultTabSizeEmu).toBe(inches(1));
  const before = getSlideXmlString(slide);
  for (const positionEmu of [NaN, Infinity, 2147483648]) {
    expect(() =>
      setParagraphTabs(shape, 0, {
        defaultTabSizeEmu: 0,
        tabStops: [{ positionEmu, alignment: 'left' }],
      }),
    ).toThrow(RangeError);
    expect(getSlideXmlString(slide)).toBe(before);
  }
  expect(() =>
    setParagraphTabs(shape, 0, {
      tabStops: [
        { positionEmu: 1, alignment: 'left' },
        { positionEmu: 1, alignment: 'right' },
      ],
    }),
  ).toThrow(RangeError);
  expect(getSlideXmlString(slide)).toBe(before);
  setParagraphTabs(shape, 0, { defaultTabSizeEmu: null });
  expect(getParagraphPropertiesEffective(pres, shape, 0).defaultTabSizeEmu).toBe(457200);
});
