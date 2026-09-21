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
describe.each(['pie', 'doughnut'] as const)('%s data labels', (kind) => {
  it('renders series names and combines label parts', async () => {
    const svg = await render({
      kind,
      categories: ['A', 'B'],
      series: [{ name: 'Sales & costs', values: [10, 30] }],
      dataLabels: {
        showValue: true,
        showCategory: true,
        showSeriesName: true,
        showPercent: true,
        numberFormat: '0.00',
        separator: ' / ',
      },
    });
    expect(svg).toContain('Sales &amp; costs / A / 10.00 / 25%');
    expect(svg).toContain('Sales &amp; costs / B / 30.00 / 75%');
  });
  it('honors series and point overrides including literal text and style', async () => {
    const hidden = {
      showValue: false,
      showCategory: false,
      showSeriesName: false,
      showPercent: false,
    };
    const svg = await render({
      kind,
      categories: ['A', 'B', 'C', 'D'],
      dataLabels: { ...hidden, showValue: true },
      series: [
        {
          name: 'Sales',
          values: [10, 20, 30, 40],
          dataLabels: { ...hidden, showCategory: true },
          pointDataLabels: [
            {
              ...hidden,
              text: 'Custom <label>',
              position: 'outEnd',
              textStyle: { color: '#123456', sizePt: 18 },
            },
            hidden,
            { ...hidden, showValue: true, numberFormat: '0.00' },
          ],
        },
      ],
    });
    expect(svg).toContain('Custom &lt;label&gt;');
    expect(svg).toContain('>30.00</text>');
    expect(svg).toContain('>D</text>');
    expect(svg).not.toContain('>B</text>');
    expect(svg).not.toContain('>20</text>');
    const labels = attrsOf(svg, 'text');
    expect(
      labels.some((label) => label.fill === '#123456' && Number(label['font-size']) === 24),
    ).toBe(true);
  });
});
