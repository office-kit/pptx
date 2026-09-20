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

describe.each(['column', 'bar'] as const)('%s fill opacity', (kind) => {
  describe.each(['clustered', 'stacked', 'percentStacked'] as const)('%s grouping', (grouping) => {
    it.each([
      [0, '0.000'],
      [0.35, '0.350'],
      [0.12345, '0.123'],
      [1, undefined],
      [undefined, undefined],
    ] as const)(
      'renders opacity %s without changing the stack',
      async (fillOpacity, expectedOpacity) => {
        const svg = await render({
          kind,
          grouping,
          categories: ['A'],
          valueAxis: { min: 0, max: grouping === 'percentStacked' ? 1 : 100 },
          series: [
            {
              name: 'Base',
              values: [20],
              color: '#010203',
              ...(fillOpacity === undefined ? {} : { fillOpacity }),
            },
            { name: 'Visible', values: [30], color: '#AABBCC' },
          ],
        });
        const rects = attrsOf(svg, 'rect');
        const base = rects.find((r) => r.fill === '#010203')!;
        const visible = rects.find((r) => r.fill === '#AABBCC')!;
        expect(base['fill-opacity']).toBe(expectedOpacity);
        expect(visible['fill-opacity']).toBeUndefined();
        if (grouping !== 'clustered') {
          const fraction = grouping === 'percentStacked' ? 0.4 : 0.2;
          if (kind === 'column') {
            expect(Number(base.height)).toBeCloseTo(446 * fraction);
            expect(Number(visible.y) + Number(visible.height)).toBeCloseTo(550 - 446 * fraction);
          } else {
            expect(Number(base.width)).toBeCloseTo(720 * fraction);
            expect(Number(visible.x)).toBeCloseTo(136 + 720 * fraction);
          }
        }
      },
    );
  });
});
