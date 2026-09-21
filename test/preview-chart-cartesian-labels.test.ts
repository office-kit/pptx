import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideChart,
  findSlideLayout,
  inches,
  loadPresentation,
  type ChartSpec,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { attrsOf } from './lib/svg-query.ts';

async function render(spec: ChartSpec) {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
}
describe.each(['column', 'bar', 'line', 'area'] as const)('%s labels', (kind) => {
  it.each(['standard', 'stacked', 'percentStacked'] as const)(
    'renders category and series names with %s grouping',
    async (grouping) => {
      const svg = await render({
        kind,
        grouping,
        categories: ['Category A'],
        categoryAxisHidden: true,
        valueAxisHidden: true,
        dataLabels: {
          showValue: false,
          showCategory: true,
          showSeriesName: true,
          showPercent: false,
          separator: ' / ',
        },
        series: [{ name: 'Sales & costs', values: [10] }],
      });
      expect(svg).toContain('Sales &amp; costs / Category A');
    },
  );
  it.each(['standard', 'stacked', 'percentStacked'] as const)(
    'honors point label visibility, literal text and style with %s grouping',
    async (grouping) => {
      const off = {
        showValue: false,
        showCategory: false,
        showSeriesName: false,
        showPercent: false,
      };
      const svg = await render({
        kind,
        grouping,
        categories: ['A', 'B'],
        categoryAxisHidden: true,
        valueAxisHidden: true,
        dataLabels: { ...off, showValue: true },
        series: [
          {
            name: 'Sales',
            dataLabels: { ...off, showValue: true },
            values: [10, 30],
            pointDataLabels: [
              { ...off, text: 'Custom <label>', textStyle: { color: '#123456', sizePt: 18 } },
              off,
            ],
          },
        ],
      });
      expect(svg).toContain('Custom &lt;label&gt;');
      expect(svg).not.toContain('>30</text>');
      expect(svg).not.toContain('>100%</text>');
      expect(
        attrsOf(svg, 'text').some(
          (label) => label.fill === '#123456' && Number(label['font-size']) === 24,
        ),
      ).toBe(true);
    },
  );
});
