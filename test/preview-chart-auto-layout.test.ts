// Automatic plot-area layout of cartesian charts, following LibreOffice's
// chart2 view (ChartView.cxx): 2% of the frame clear on each side, a
// category-label row sized to the label font, and line-chart labels kept 1 mm
// clear of their markers.

import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideChart,
  findSlideLayout,
  inches,
  loadPresentation,
  type ChartSpec,
  type ComboChartSpec,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { attrsOf } from './lib/svg-query.ts';

const PX_PER_PT = 96 / 72;
const PX_PER_MM = 96 / 25.4;
// The chart frame: 8 × 5 in at (1 in, 1 in).
const FRAME = { x: 96, y: 96, w: 768, h: 480 };

async function render(spec: ChartSpec) {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
}

const column: ComboChartSpec = {
  kind: 'column',
  categories: ['A', 'B'],
  series: [{ name: 'Sessions', values: [60, 80], color: '#123456' }],
  valueAxis: { min: 0, max: 100 },
  plotAreaFill: '#FFFFCC',
};

const plotOf = (svg: string) => {
  const plot = attrsOf(svg, 'rect').find((r) => r.fill === '#FFFFCC')!;
  const [x, y, w, h] = [plot.x, plot.y, plot.width, plot.height].map(Number) as [
    number,
    number,
    number,
    number,
  ];
  return { x, y, w, h, bottom: y + h };
};

describe('automatic cartesian chart layout', () => {
  it('keeps 2% of the frame clear on each side', async () => {
    const plot = plotOf(await render(column));

    expect(plot.x).toBeCloseTo(FRAME.x + FRAME.w * 0.02 + 32, 1);
    expect(plot.y).toBeCloseTo(FRAME.y + FRAME.h * 0.02, 1);
    expect(plot.x + plot.w).toBeCloseTo(FRAME.x + FRAME.w * 0.98, 1);
  });

  it('gives larger category labels a taller row under the axis', async () => {
    const small = plotOf(await render(column));
    const large = plotOf(await render({ ...column, categoryAxisLabelStyle: { sizePt: 18 } }));

    expect(small.bottom - large.bottom).toBeCloseTo(((18 - 10) * PX_PER_PT * 1.2) / 2, 1);
  });

  it('keeps a line label above its point 1 mm clear of the marker', async () => {
    const labelPt = 12;
    const markerPt = 7;
    const svg = await render({
      kind: 'line',
      categories: ['A', 'B'],
      series: [{ name: 'Rate', values: [50, 50], markerSymbol: 'circle', markerSizePt: markerPt }],
      valueAxis: { min: 0, max: 100 },
      dataLabels: {
        showValue: true,
        showCategory: false,
        showSeriesName: false,
        showPercent: false,
        position: 't',
        textStyle: { sizePt: labelPt },
      },
      plotAreaFill: '#FFFFCC',
    });
    const plot = plotOf(svg);
    const pointY = plot.y + plot.h / 2;
    const label = [...svg.matchAll(/<text x="[\d.]+" y="([\d.]+)"[^>]*>50<\/text>/g)];
    const descent = labelPt * PX_PER_PT * 0.269;

    expect(label).toHaveLength(2);
    expect(Number(label[0]![1])).toBeCloseTo(
      pointY - (markerPt * PX_PER_PT) / 2 - PX_PER_MM - descent,
      1,
    );
  });
});
