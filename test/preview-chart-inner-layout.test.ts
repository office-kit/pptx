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
