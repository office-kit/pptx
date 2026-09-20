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
  const layout = findSlideLayout(pres, 'Blank')!;
  const slide = addSlide(pres, { layout });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
}
const base: ChartSpec = {
  kind: 'column',
  categories: ['A'],
  series: [{ name: 'S', values: [60], color: '#123456' }],
  valueAxis: { min: 0, max: 100 },
  plotAreaFill: '#FFFFCC',
};

describe('manual inner chart layout', () => {
  it.each([false, true])(
    'uses frame-relative geometry with title and legend: %s',
    async (decorated) => {
      const svg = await render({
        ...base,
        ...(decorated ? { title: 'Title', legend: { position: 'b' as const } } : {}),
        plotAreaLayout: { target: 'inner', x: 0.2, y: 0.2, w: 0.6, h: 0.6 },
      });
      const plot = attrsOf(svg, 'rect').find((r) => r.fill === '#FFFFCC')!;
      expect(Number(plot.x)).toBeCloseTo(249.6);
      expect(Number(plot.y)).toBeCloseTo(192);
      expect(Number(plot.width)).toBeCloseTo(460.8);
      expect(Number(plot.height)).toBeCloseTo(288);
      const bar = attrsOf(svg, 'rect').find((r) => r.fill === '#123456')!;
      expect(Number(bar.x) + Number(bar.width) / 2).toBeCloseTo(480);
      expect(Number(bar.y)).toBeCloseTo(307.2);
      expect(Number(bar.height)).toBeCloseTo(172.8);
    },
  );
  it.each([
    { name: 'negative sizes', x: 0.5, y: 0.5, w: -0.3, h: -0.3, expected: [0.5, 0.5, 0, 0] },
    {
      name: 'right and bottom overflow',
      x: 0.5,
      y: 0.5,
      w: 0.6,
      h: 0.6,
      expected: [0.5, 0.5, 0.5, 0.5],
    },
    { name: 'left and top overflow', x: -0.2, y: -0.2, w: 0.6, h: 0.6, expected: [0, 0, 0.4, 0.4] },
    { name: 'origin outside frame', x: 1.2, y: 1.2, w: 0.6, h: 0.6, expected: [1, 1, 0, 0] },
  ])('clamps $name to the chart frame', async ({ x, y, w, h, expected }) => {
    const svg = await render({
      ...base,
      plotAreaLayout: { target: 'inner', x, y, w, h },
    });
    const plot = attrsOf(svg, 'rect').find((r) => r.fill === '#FFFFCC')!;
    const pxPerInch = 96;
    const frame = { x: pxPerInch, y: pxPerInch, w: 8 * pxPerInch, h: 5 * pxPerInch };
    expect(Number(plot.x)).toBeCloseTo(frame.x + expected[0]! * frame.w);
    expect(Number(plot.y)).toBeCloseTo(frame.y + expected[1]! * frame.h);
    expect(Number(plot.width)).toBeCloseTo(expected[2]! * frame.w);
    expect(Number(plot.height)).toBeCloseTo(expected[3]! * frame.h);
    for (const rect of attrsOf(svg, 'rect')) {
      expect(Number(rect.width)).toBeGreaterThanOrEqual(0);
      expect(Number(rect.height)).toBeGreaterThanOrEqual(0);
    }
  });
  it.each([undefined, 'outer' as const])(
    'retains automatic gutters for %s layout',
    async (target) => {
      const svg = await render({
        ...base,
        ...(target ? { plotAreaLayout: { target, x: 0.2, y: 0.2, w: 0.6, h: 0.6 } } : {}),
      });
      const plot = attrsOf(svg, 'rect').find((r) => r.fill === '#FFFFCC')!;
      expect([plot.x, plot.y, plot.width, plot.height].map(Number)).toEqual([136, 104, 720, 446]);
    },
  );
});
