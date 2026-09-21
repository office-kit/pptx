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
describe.each(['stacked', 'percentStacked'] as const)('%s line markers', (grouping) => {
  it('places explicit markers on each cumulative line point', async () => {
    const svg = await render({
      kind: 'line',
      grouping,
      categories: ['A', 'B', 'C'],
      series: [
        { name: 'First', values: [10, 20, 15], markerSymbol: 'none' },
        {
          name: 'Second',
          values: [20, 20, 20],
          color: '#123456',
          markerSymbol: 'circle',
          markerSizePt: 12,
        },
      ],
    });
    const circles = attrsOf(svg, 'circle').filter(
      (circle) => circle.fill === '#123456' && circle.r === '6.00',
    );
    expect(circles).toHaveLength(3);
    const path = attrsOf(svg, 'path').find((path) => path.stroke === '#123456');
    for (const circle of circles) expect(path?.d).toContain(`${circle.cx},${circle.cy}`);
  });
});

describe('smoothed line gaps', () => {
  it.each(['gap', 'span', 'zero'] as const)(
    'honors %s for missing values',
    async (dispBlanksAs) => {
      const svg = await render({
        kind: 'line',
        categories: ['A', 'B', 'C', 'D', 'E', 'F', 'G'],
        dispBlanksAs,
        series: [
          {
            name: 'S',
            values: [10, 20, 15, null, 25, 10, 20],
            color: '#123456',
            smooth: true,
            markerSymbol: 'circle',
          },
        ],
      });
      const path = attrsOf(svg, 'path').find((path) => path.stroke === '#123456')?.d ?? '';
      expect(path.match(/M/g)).toHaveLength(dispBlanksAs === 'gap' ? 2 : 1);
      expect(path.match(/C/g)).toHaveLength(
        dispBlanksAs === 'gap' ? 4 : dispBlanksAs === 'zero' ? 6 : 5,
      );
      expect(attrsOf(svg, 'circle')).toHaveLength(dispBlanksAs === 'zero' ? 7 : 6);
    },
  );
});

it('retains isolated points and straight two-point segments around consecutive gaps', async () => {
  const svg = await render({
    kind: 'line',
    categories: ['A', 'B', 'C', 'D', 'E', 'F', 'G'],
    series: [
      {
        name: 'S',
        values: [null, 10, null, null, 20, 30, null],
        color: '#123456',
        smooth: true,
        markerSymbol: 'circle',
      },
    ],
  });
  const path = attrsOf(svg, 'path').find((path) => path.stroke === '#123456')?.d ?? '';
  expect(path.match(/M/g)).toHaveLength(2);
  expect(path.match(/L/g)).toHaveLength(1);
  expect(path).not.toContain('C');
  expect(attrsOf(svg, 'circle')).toHaveLength(3);
});
