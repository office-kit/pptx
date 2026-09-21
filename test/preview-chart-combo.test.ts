// Combo chart rendering — a line overlay on a column chart with a
// secondary (right-hand) value axis. Asserts on the emitted SVG:
//
//   - both the bars and the overlay polyline are painted
//   - the secondary axis ticks sit on the plot's RIGHT edge and are
//     scaled to the secondary series (not squashed by the primary range)

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
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

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const renderChart = async (spec: ChartSpec): Promise<string> => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
};

describe('preview: combo charts', () => {
  it('renders bars + line overlay with a right-hand secondary axis', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['速い', '普通', '遅い', '非常に遅い'],
      series: [
        { name: '件数', values: [92, 118, 54, 17], color: '#4472C4' },
        {
          name: '平均スコア',
          values: [4.5, 3.8, 2.9, 2.1],
          color: '#ED7D31',
          chartKind: 'line',
          secondaryAxis: true,
        },
      ],
    });

    // Bars from the column group…
    expect((svg.match(/<rect[^>]*fill="#4472C4"/g) ?? []).length).toBeGreaterThanOrEqual(4);
    // …and the line overlay's path in the series color.
    expect(svg).toMatch(/<path[^>]*stroke="#ED7D31"/);

    // Secondary axis ticks: labels anchored `start` (right edge) exist,
    // and the secondary scale (0..5-ish) appears — a squashed overlay
    // on the primary 0..120 scale would never emit a "4" tick label.
    expect(svg).toMatch(/<text[^>]*text-anchor="start"[^>]*>\s*[45]\s*<\/text>/);
  });

  it('keeps the single-kind path untouched when no combo fields are set', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'v', values: [1, 2], color: '#4472C4' }],
    });

    expect((svg.match(/<rect[^>]*fill="#4472C4"/g) ?? []).length).toBeGreaterThanOrEqual(2);
    expect(svg).not.toMatch(/text-anchor="start"[^>]*dominant-baseline="middle"/);
  });
});
