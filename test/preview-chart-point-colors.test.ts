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

describe.each(['column', 'bar'] as const)('%s point colors', (kind) => {
  it.each(['stacked', 'percentStacked', 'clustered'] as const)(
    'preserves per-point overrides and series fallbacks in %s',
    async (grouping) => {
      const svg = await render({
        kind,
        grouping,
        categories: ['A', 'B', 'C'],
        series: [
          {
            name: 'Highlighted',
            values: [20, 30, grouping === 'percentStacked' ? 40 : -40],
            color: '#ABCDEF',
            pointColors: [null, '#E5481F', null],
          },
          { name: 'Other', values: [10, 15, 20], color: '#123456' },
        ],
      });
      const bars = attrsOf(svg, 'rect').filter((r) =>
        ['#ABCDEF', '#E5481F', '#123456'].includes(r.fill!),
      );
      expect(bars.map((r) => r.fill)).toEqual([
        '#ABCDEF',
        '#123456',
        '#E5481F',
        '#123456',
        '#ABCDEF',
        '#123456',
      ]);
      expect(bars.every((r) => Number(r.width) > 0 && Number(r.height) > 0)).toBe(true);
    },
  );
  it('uses a point override before single-series varying colors', async () => {
    const svg = await render({
      kind,
      grouping: 'stacked',
      varyColors: true,
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [20, 30], pointColors: ['#E5481F', null] }],
    });
    const bars = attrsOf(svg, 'rect').filter(
      (r) => Number(r.width) > 0 && Number(r.height) > 0 && r.fill !== 'none',
    );
    expect(
      bars.filter((r) => r.fill === '#E5481F' || r.fill === '#C0504D').map((r) => r.fill),
    ).toEqual(['#E5481F', '#C0504D']);
  });
});
