import { expect, it } from 'vitest';
import {
  createPresentation,
  addBlankSlide,
  addSlideTextBox,
  addSlideTable,
  getTableCells,
  getSlideShapes,
  getSlides,
  getSlideSize,
  inches,
  loadPresentation,
  savePresentation,
  setParagraphTabs,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';

it.each(['shape', 'cell'] as const)(
  'renders saved %s tab stops and default spacing at their authored positions',
  async (kind) => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const bounds = { x: inches(1), y: inches(1), w: inches(5), h: inches(2) };
    const shape =
      kind === 'shape'
        ? addSlideTextBox(slide, { ...bounds, text: 'A\tB' })
        : addSlideTable(slide, { ...bounds, rows: [['A\tB']] });
    const target = kind === 'shape' ? shape : getTableCells(shape)[0]![0]!;
    const options = {
      textLayout: 'svg' as const,
      measureText: (text: string) => ({ widthPx: text.length * 10 }),
    };
    const gap = (svg: string) => Number(svg.match(/<tspan dx="([^"]+)">&#8203;/)?.[1]);
    setParagraphTabs(target, 0, { tabStops: [{ positionEmu: inches(1), alignment: 'left' }] });
    const first = renderSlideToSvg(pres, slide, options);
    setParagraphTabs(target, 0, { tabStops: [{ positionEmu: inches(2), alignment: 'left' }] });
    const second = renderSlideToSvg(pres, slide, options);
    expect(gap(second) - gap(first)).toBe(96);
    const loaded = await loadPresentation(await savePresentation(pres));
    const loadedSlide = getSlides(loaded)[0]!;
    expect(gap(renderSlideToSvg(loaded, loadedSlide, options))).toBe(gap(second));
    const loadedShape = getSlideShapes(loadedSlide)[0]!;
    const loadedTarget = kind === 'shape' ? loadedShape : getTableCells(loadedShape)[0]![0]!;
    setParagraphTabs(loadedTarget, 0, { tabStops: [], defaultTabSizeEmu: inches(1) });
    expect(gap(renderSlideToSvg(loaded, loadedSlide, options))).toBe(gap(first));
  },
);

it('moves rasterized glyphs by the tab stop displacement', () => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const shape = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(5),
    h: inches(2),
    text: '\tTEST',
  });
  const minima = [1, 2].map((position) => {
    setParagraphTabs(shape, 0, {
      tabStops: [{ positionEmu: inches(position), alignment: 'left' }],
    });
    const { image } = renderSlideToRgba(pres, slide, {
      width: Math.round(getSlideSize(pres)!.width / 9525),
    });
    let minimum = image.width;
    for (let i = 0; i < image.data.length; i += 4) {
      if (
        image.data[i + 3]! > 128 &&
        image.data[i]! < 128 &&
        image.data[i + 1]! < 128 &&
        image.data[i + 2]! < 128
      ) {
        minimum = Math.min(minimum, (i / 4) % image.width);
      }
    }
    expect(minimum).toBeLessThan(image.width);
    return minimum;
  });
  expect(minima[1]! - minima[0]!).toBe(96);
});
