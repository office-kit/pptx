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

describe.each(['stacked', 'percentStacked'] as const)('%s label placement', (grouping) => {
  describe.each(['column', 'bar'] as const)('%s stacked label placement', (kind) => {
    it.each(['ctr', 'inEnd', 'outEnd', 'inBase'] as const)(
      'places %s labels relative to each signed segment',
      async (position) => {
        const svg = await render({
          kind,
          grouping,
          categories: ['Positive', 'Negative'],
          categoryAxisHidden: true,
          valueAxisHidden: true,
          dataLabels: {
            showValue: true,
            showCategory: false,
            showSeriesName: false,
            showPercent: false,
            position,
            textStyle: { color: '#123456' },
          },
          series: [
            { name: 'First', values: [10, grouping === 'stacked' ? -10 : 10], color: '#ABCDEF' },
            { name: 'Second', values: [20, grouping === 'stacked' ? -20 : 20], color: '#FEDCBA' },
          ],
        });
        const bars = attrsOf(svg, 'rect').filter(
          (rect) => rect.fill === '#ABCDEF' || rect.fill === '#FEDCBA',
        );
        const labels = attrsOf(svg, 'text').filter((label) => label.fill === '#123456');
        expect(labels).toHaveLength(4);
        for (let i = 0; i < 4; i++) {
          const bar = bars[i]!;
          const label = labels[i]!;
          const positive = grouping === 'percentStacked' || i < 2;
          if (kind === 'column') {
            const y = Number(bar.y),
              h = Number(bar.height);
            const expected =
              position === 'ctr'
                ? y + h / 2 + 3
                : position === 'inEnd'
                  ? positive
                    ? y + 9
                    : y + h - 3
                  : position === 'inBase'
                    ? positive
                      ? y + h - 3
                      : y + 9
                    : positive
                      ? y - 2
                      : y + h + 9;
            expect(Number(label.y)).toBeCloseTo(expected, 1);
          } else {
            const x = Number(bar.x),
              w = Number(bar.width);
            const expected =
              position === 'ctr'
                ? x + w / 2
                : position === 'inEnd'
                  ? positive
                    ? x + w - 4
                    : x + 4
                  : position === 'inBase'
                    ? positive
                      ? x + 4
                      : x + w - 4
                    : positive
                      ? x + w + 2
                      : x - 2;
            expect(Number(label.x)).toBeCloseTo(expected, 1);
          }
        }
      },
    );
  });
});

describe.each(['column', 'bar', 'line', 'area'] as const)('%s percent-stacked values', (kind) => {
  it('formats original values with chart, series and point formats', async () => {
    const visible = {
      showValue: true,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
    };
    const svg = await render({
      kind,
      grouping: 'percentStacked',
      categories: ['A', 'B'],
      categoryAxisHidden: true,
      valueAxisHidden: true,
      dataLabels: { ...visible, numberFormat: '0.00' },
      series: [
        { name: 'First', values: [10, 20] },
        {
          name: 'Second',
          values: [30, 40],
          dataLabels: { ...visible, numberFormat: '0.0' },
          pointDataLabels: [{ ...visible, numberFormat: '0.000' }],
        },
      ],
    });
    for (const label of ['10.00', '20.00', '30.000', '40.0'])
      expect(svg).toContain(`>${label}</text>`);
  });
  it('shows original values when no label format is specified', async () => {
    const svg = await render({
      kind,
      grouping: 'percentStacked',
      categories: ['A'],
      categoryAxisHidden: true,
      valueAxisHidden: true,
      dataLabels: {
        showValue: true,
        showCategory: false,
        showSeriesName: false,
        showPercent: false,
      },
      series: [
        { name: 'First', values: [10] },
        { name: 'Second', values: [30] },
      ],
    });
    expect(svg).toContain('>10</text>');
    expect(svg).toContain('>30</text>');
  });
});
