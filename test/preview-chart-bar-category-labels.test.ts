// Bar-chart category labels sit left of the plot. Presentation apps narrow the
// plot so the labels fit instead of ellipsizing them, and wrap a label only
// when it would take more than half the chart's width.

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

// The chart frame below starts 1in (96px) from the slide's left edge.
const CHART_LEFT_PX = 96;
const CHART_WIDTH_PX = 6 * 96;
const BAR_COLOR = '#2E75B6';

async function render(spec: ChartSpec) {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(6), h: inches(4), spec });
  return renderSlideToSvg(pres, slide);
}

const barChart = (categories: string[]): ComboChartSpec => ({
  kind: 'bar',
  categories,
  series: [{ name: 'n', values: categories.map((_, i) => (i + 1) * 10), color: BAR_COLOR }],
});

const barLeft = (svg: string): number =>
  Math.min(
    ...attrsOf(svg, 'rect')
      .filter((r) => r.fill === BAR_COLOR)
      .map((r) => Number(r.x)),
  );

describe('preview: bar-chart category labels', () => {
  it('narrows the plot to fit long labels instead of ellipsizing them', async () => {
    const labels = ['本人確認が必要な手続き（22%）', 'ナレッジに回答が無い（34%）'];
    const short = await render(barChart(['A', 'B']));
    const long = await render(barChart(labels));

    for (const label of labels) expect(long).toContain(`>${label}</text>`);
    expect(long).not.toContain('…');
    expect(barLeft(long)).toBeGreaterThan(barLeft(short));
  });

  it('wraps a label wider than half the chart, keeping closing brackets on the line', async () => {
    const label =
      'お問い合わせ窓口で本人確認が必要になる手続きのうち書類の再提出を求めたもの（22%）';
    const svg = await render(barChart([label, '短い']));

    const lines = [...svg.matchAll(/<tspan[^>]*>([^<]*)<\/tspan>/g)].map((m) => m[1]!);
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.join('')).toBe(label);
    expect(lines.some((line) => /^[）%]/u.test(line))).toBe(false);
    expect(barLeft(svg) - CHART_LEFT_PX).toBeLessThanOrEqual(CHART_WIDTH_PX / 2 + 40);
  });

  it('wraps Latin labels at spaces', async () => {
    const label =
      'Customers who asked for a human agent after the assistant answered the wrong question';
    const svg = await render(barChart([label]));

    const lines = [...svg.matchAll(/<tspan[^>]*>([^<]*)<\/tspan>/g)].map((m) => m[1]!);
    expect(lines.length).toBeGreaterThan(1);
    expect(lines.join(' ')).toBe(label);
  });

  it('puts the value-axis title under the horizontal value axis', async () => {
    const svg = await render({ ...barChart(['A', 'B']), valueAxisTitle: '件数' });

    const title = /<text([^>]*)>件数<\/text>/.exec(svg)?.[1];
    expect(title).toBeDefined();
    expect(title).not.toContain('rotate(');
    const plotBottom = Math.max(
      ...attrsOf(svg, 'rect')
        .filter((r) => r.fill === BAR_COLOR)
        .map((r) => Number(r.y) + Number(r.height)),
    );
    expect(Number(/\by="([\d.]+)"/.exec(title!)?.[1])).toBeGreaterThan(plotBottom);
  });
});
