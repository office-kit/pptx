import { describe, expect, it } from 'vitest';
import * as api from '../src/api/index.ts';
import { Chart, Presentation, Slide, Text, compile } from '../packages/dsl/src/index.ts';

describe('Presentation fonts', () => {
  it('sets the theme typefaces that charts and unformatted text inherit', async () => {
    const pres = await compile(
      Presentation({
        fonts: { minorLatin: 'Arial', majorLatin: 'Arial' },
        children: Slide({
          children: [
            Text({ x: 1, y: 1, width: 4, height: 1, children: '本文' }),
            Chart({
              x: 1,
              y: 2,
              width: 6,
              height: 4,
              spec: {
                kind: 'column',
                categories: ['4 月', '5 月'],
                series: [{ name: '件数', values: [1, 2] }],
              },
            }),
          ],
        }),
      }),
    );

    expect(api.getPresentationFonts(pres)).toMatchObject({
      minorLatin: 'Arial',
      majorLatin: 'Arial',
      // East Asian text keeps the theme's own East Asian and script fonts.
      minorEastAsian: null,
    });
    const slide = api.getSlides(pres)[0]!;
    const chart = api.getSlideShapes(slide).find((shape) => api.getShapeChartSpec(shape))!;
    // The chart names no typeface, so it draws with the theme's.
    expect(api.getShapeChartSpec(chart)?.textStyle?.font).toBeUndefined();
  });

  it('leaves the theme fonts alone when omitted', async () => {
    const plain = await compile(Presentation({ children: Slide({}) }));
    expect(api.getPresentationFonts(plain)?.minorLatin).toBe(
      api.getPresentationFonts(api.createPresentation())?.minorLatin,
    );
  });
});
