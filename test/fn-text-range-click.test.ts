import { describe, expect, it } from 'vitest';
import * as pptx from '../src/api/index.ts';

function fixture() {
  const presentation = pptx.createPresentation();
  const slide = pptx.addBlankSlide(presentation);
  const shape = pptx.addSlideTextBox(slide, {
    x: pptx.inches(1),
    y: pptx.inches(1),
    w: pptx.inches(5),
    h: pptx.inches(2),
    text: '',
  });
  pptx.setShapeParagraphs(shape, [
    { runs: [{ text: 'A😀B', format: { bold: true } }] },
    { runs: [{ text: '日本語', format: { italic: true } }] },
  ]);
  return { presentation, slide, shape };
}
describe('text-range click actions', () => {
  it('links a Unicode range across paragraphs while preserving surrounding links and formatting', async () => {
    const { presentation, shape } = fixture();
    pptx.setShapeRunHyperlink(shape, 0, 0, 'https://previous.example/', 'Previous');
    pptx.setShapeClickAction(shape, { kind: 'firstSlide' }, 'Object');
    pptx.setShapeTextRangeClickAction(
      shape,
      1,
      7,
      { kind: 'url', url: 'https://example.com/?a=1&b=2' },
      '資料 <詳細>',
    );
    const loaded = await pptx.loadPresentation(await pptx.savePresentation(presentation));
    const result = pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!;
    expect(pptx.getShapeText(result)).toBe('A😀B\n日本語');
    expect(pptx.getShapeParagraphElements(result, 0)).toMatchObject([
      { text: 'A', format: { bold: true } },
      { text: '😀B', format: { bold: true } },
    ]);
    expect(pptx.getShapeParagraphElements(result, 1)).toMatchObject([
      { text: '日本', format: { italic: true } },
      { text: '語', format: { italic: true } },
    ]);
    expect(pptx.getShapeRunHyperlink(result, 0, 0)).toBe('https://previous.example/');
    expect(pptx.getShapeRunHyperlinkTooltip(result, 0, 0)).toBe('Previous');
    for (const [p, r] of [
      [0, 1],
      [1, 0],
    ] as const) {
      expect(pptx.getShapeRunClickAction(result, p, r)).toEqual({
        kind: 'url',
        url: 'https://example.com/?a=1&b=2',
      });
      expect(pptx.getShapeRunHyperlinkTooltip(result, p, r)).toBe('資料 <詳細>');
    }
    expect(pptx.getShapeRunClickAction(result, 1, 1)).toBeNull();
    expect(pptx.getShapeClickAction(result)).toEqual({ kind: 'firstSlide' });
    pptx.setShapeTextRangeClickAction(result, 1, 3, null);
    expect(pptx.getShapeRunHyperlink(result, 0, 1)).toBeNull();
    expect(pptx.getShapeRunHyperlink(result, 0, 2)).toBe('https://example.com/?a=1&b=2');
    expect(pptx.getShapeClickActionTooltip(result)).toBe('Object');
  });
  it('round-trips internal and preset links and rejects invalid ranges and foreign slides atomically', async () => {
    const { presentation, slide, shape } = fixture();
    const target = pptx.addBlankSlide(presentation);
    pptx.setShapeTextRangeClickAction(shape, 0, 1, { kind: 'slide', slide: target }, 'Go');
    const loaded = await pptx.loadPresentation(await pptx.savePresentation(presentation));
    const action = pptx.getShapeRunClickAction(
      pptx.getSlideShapes(pptx.getSlides(loaded)[0]!)[0]!,
      0,
      0,
    );
    expect(action?.kind).toBe('slide');
    if (action?.kind === 'slide')
      expect(pptx.getSlidePartName(action.slide)).toBe(pptx.getSlidePartName(target));
    for (const kind of ['nextSlide', 'prevSlide', 'firstSlide', 'lastSlide'] as const) {
      pptx.setShapeTextRangeClickAction(shape, 1, 3, { kind });
      expect(pptx.getShapeRunClickAction(shape, 0, 1)).toEqual({ kind });
    }
    const before = pptx.getSlideXmlString(slide);
    for (const [start, end] of [
      [2, 3],
      [-1, 1],
      [0, 99],
      [3, 1],
    ])
      expect(() => pptx.setShapeTextRangeClickAction(shape, start!, end!, null)).toThrow();
    const foreign = pptx.addBlankSlide(pptx.createPresentation());
    expect(() =>
      pptx.setShapeTextRangeClickAction(shape, 0, 1, { kind: 'slide', slide: foreign }),
    ).toThrow('target slide must belong');
    pptx.setShapeTextRangeClickAction(shape, 0, 0, { kind: 'url', url: 'https://unused.example/' });
    expect(pptx.getSlideXmlString(slide)).toBe(before);
  });
});
