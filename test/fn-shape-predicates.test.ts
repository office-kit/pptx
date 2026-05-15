// isChartShape / isTableShape — type narrowing for graphic-frame
// shapes. getShapeKind returns 'graphicFrame' for charts, tables, and
// SmartArt; these predicates distinguish charts from tables.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideChart,
  addSlideShape,
  addSlideTable,
  getShapeKind,
  getSlides,
  inches,
  isChartShape,
  isTableShape,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

describe('fn API: isTableShape / isChartShape', () => {
  it('classifies tables and charts independently', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const table = addSlideTable(slide, {
      x: inches(0), y: inches(0), w: inches(4), h: inches(2),
      rows: [['a', 'b']],
    });
    const chart = addSlideChart(slide, {
      x: inches(0), y: inches(3), w: inches(4), h: inches(3),
      spec: {
        kind: 'bar',
        categories: ['Q1'],
        series: [{ name: 'r', values: [1] }],
      },
    });

    expect(isTableShape(table)).toBe(true);
    expect(isChartShape(table)).toBe(false);

    expect(isChartShape(chart)).toBe(true);
    expect(isTableShape(chart)).toBe(false);

    // Both are graphic frames; the predicates discriminate.
    expect(getShapeKind(table)).toBe('graphicFrame');
    expect(getShapeKind(chart)).toBe('graphicFrame');
  });

  it('returns false for non-graphic-frame shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const rect = addSlideShape(slide, {
      preset: 'rect', x: inches(0), y: inches(0), w: inches(1), h: inches(1),
    });
    expect(isTableShape(rect)).toBe(false);
    expect(isChartShape(rect)).toBe(false);
  });
});
