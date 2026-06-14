// Structural regression tests for the chart-rendering fidelity fixes that
// were validated against PowerPoint PDF exports (see the chart-gallery sample).
// These assert on the emitted SVG so they run in CI without PowerPoint:
//
//   - no invented chart-area border (PowerPoint draws none unless authored)
//   - value + category axis spines and major tick marks are drawn
//   - bar charts order categories bottom-to-top
//   - stacked charts use the authored series color and scale the value axis
//     to the per-category stacked total
//   - area fills are opaque and draw no markers
//   - line markers follow the "Line with Markers" subtype + PowerPoint's
//     automatic marker-symbol rotation (diamond, square, triangle, x, …)

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
import { countTags } from './lib/svg-query.ts';

const fixturePath = fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const renderChart = async (spec: ChartSpec): Promise<string> => {
  const pres = await loadPresentation(await readFile(fixturePath));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('Blank layout not found');
  const slide = addSlide(pres, { layout });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
};

// Count `<line>` elements painted in the default near-black axis color.
const blackLines = (svg: string): number => (svg.match(/<line[^>]*stroke="#000000"/g) ?? []).length;

// The y of the first `<text>` whose content is exactly `label`.
const labelY = (svg: string, label: string): number => {
  const m = svg.match(new RegExp(`<text([^>]*)>${label}</text>`));
  if (!m) throw new Error(`no <text> with content "${label}"`);
  const y = m[1]!.match(/\by="([\d.]+)"/);
  if (!y) throw new Error(`<text>${label}</text> has no y`);
  return Number(y[1]);
};

// Numeric value-axis tick labels (category labels here are letters, so the
// numeric ones are the axis ticks).
const numericLabels = (svg: string): number[] =>
  [...svg.matchAll(/<text[^>]*>(\d+)<\/text>/g)].map((m) => Number(m[1]));

describe('chart fidelity vs PowerPoint', () => {
  it('draws no chart-area border by default', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
    });
    // The chart-area backdrop is white-filled with no stroke.
    expect(svg).toMatch(/<rect[^>]*fill="#FFFFFF"[^>]*stroke="none"/);
  });

  it('honors an authored chart-area border', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
      chartAreaStrokeColor: '#FF0000',
    });
    expect(svg).toContain('stroke="#FF0000"');
  });

  it('draws value + category axis spines and tick marks for a column chart', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['A', 'B', 'C', 'D'],
      series: [{ name: 'S', values: [1, 2, 3, 4] }],
    });
    // 2 spines + 5 category-boundary ticks + value-axis ticks → comfortably > 6.
    expect(blackLines(svg)).toBeGreaterThan(6);
  });

  it('does not draw value-axis gridlines unless authored', async () => {
    const svg = await renderChart({
      kind: 'column',
      categories: ['A', 'B'],
      series: [{ name: 'S', values: [1, 2] }],
    });
    // The old gridline color must not appear for an unauthored axis.
    expect(svg).not.toContain('stroke="#E5E7EB"');
  });

  it('orders bar-chart categories bottom-to-top', async () => {
    const svg = await renderChart({
      kind: 'bar',
      categories: ['Low', 'High'],
      series: [{ name: 'S', values: [1, 2] }],
    });
    // 'Low' (category 0) sits at the bottom → larger y than 'High'.
    expect(labelY(svg, 'Low')).toBeGreaterThan(labelY(svg, 'High'));
  });

  it('paints stacked-column segments in the authored series colors', async () => {
    const svg = await renderChart({
      kind: 'column',
      grouping: 'stacked',
      categories: ['A'],
      series: [
        { name: 'P1', values: [10], color: '#FF0000' },
        { name: 'P2', values: [20], color: '#00FF00' },
      ],
    });
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain('fill="#00FF00"');
  });

  it('scales a stacked value axis to the per-category total', async () => {
    const svg = await renderChart({
      kind: 'column',
      grouping: 'stacked',
      categories: ['A'],
      series: [
        { name: 'P1', values: [10] },
        { name: 'P2', values: [20] },
      ],
    });
    // Axis must reach the stacked total (30) — not the largest single value (20).
    expect(Math.max(...numericLabels(svg))).toBeGreaterThanOrEqual(30);
  });

  it('fills area charts opaquely and draws no markers', async () => {
    const svg = await renderChart({
      kind: 'area',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S', values: [1, 2, 3] }],
    });
    expect(svg).not.toContain('fill-opacity');
    // Area charts plot the fill + line as <path>; no marker glyphs.
    expect(countTags(svg, 'polygon')).toBe(0);
    expect(countTags(svg, 'circle')).toBe(0);
  });

  it('draws diamond markers for a default ("with markers") line chart', async () => {
    const svg = await renderChart({
      kind: 'line',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S', values: [1, 2, 3] }],
    });
    // Series 0's automatic marker is a diamond → one <polygon> per point.
    expect(countTags(svg, 'polygon')).toBeGreaterThanOrEqual(3);
    expect(countTags(svg, 'circle')).toBe(0);
  });

  it('draws no markers for the plain "Line" subtype (lineMarkers: false)', async () => {
    const svg = await renderChart({
      kind: 'line',
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S', values: [1, 2, 3] }],
      lineMarkers: false,
    });
    expect(countTags(svg, 'polygon')).toBe(0);
    expect(countTags(svg, 'circle')).toBe(0);
    // The connecting line is still drawn.
    expect(countTags(svg, 'path')).toBeGreaterThanOrEqual(1);
  });

  it('uses the automatic marker rotation across series (not all circles)', async () => {
    const svg = await renderChart({
      kind: 'line',
      categories: ['A', 'B'],
      series: [
        { name: '1', values: [1, 2] },
        { name: '2', values: [2, 1] },
        { name: '3', values: [1, 3] },
        { name: '4', values: [3, 1] },
      ],
    });
    // diamond (series 0) + triangle (series 2) → polygons; square (series 1)
    // → rects; none render as circles.
    expect(countTags(svg, 'circle')).toBe(0);
    expect(countTags(svg, 'polygon')).toBeGreaterThanOrEqual(4);
  });
});
