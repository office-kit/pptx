// Absent `<c:majorTickMark>` / `<c:minorTickMark>` mean `cross` per the
// schema, so the builder always writes the tick marks it intends.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  type ChartSpec,
  addSlideChart,
  getSlides,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const chartXml = async (spec: ChartSpec): Promise<string> => {
  const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
  addSlideChart(getSlides(pres)[0]!, {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(6),
    h: inches(4),
    spec,
  });
  const saved = await loadPresentation(await savePresentation(pres));
  return new TextDecoder().decode(readPackagePart(saved, '/ppt/charts/chart1.xml')!);
};

const axis = (xml: string, tag: 'catAx' | 'valAx' | 'serAx'): string =>
  xml.slice(xml.indexOf(`<c:${tag}>`), xml.indexOf(`</c:${tag}>`));

const ticks = (axisXml: string): readonly [string | undefined, string | undefined] => [
  /<c:majorTickMark val="(\w+)"\/>/.exec(axisXml)?.[1],
  /<c:minorTickMark val="(\w+)"\/>/.exec(axisXml)?.[1],
];

const series = [{ name: 'S', values: [1, 2] }];

describe('chart builder: tick mark defaults', () => {
  it.each(['column', 'bar'] as const)('writes no category ticks on a %s chart', async (kind) => {
    const xml = await chartXml({ kind, categories: ['A', 'B'], series });
    expect(ticks(axis(xml, 'catAx'))).toEqual(['none', 'none']);
    expect(ticks(axis(xml, 'valAx'))).toEqual(['out', 'none']);
  });

  it.each(['line', 'area'] as const)(
    'writes outward category ticks on a %s chart',
    async (kind) => {
      const xml = await chartXml({ kind, categories: ['A', 'B'], series });
      expect(ticks(axis(xml, 'catAx'))).toEqual(['out', 'none']);
      expect(ticks(axis(xml, 'valAx'))).toEqual(['out', 'none']);
    },
  );

  it('writes outward ticks on both value axes of a scatter chart', async () => {
    const xml = await chartXml({
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [1, 2], xValues: [3, 4] }],
    });
    const [x, y] = xml.split('</c:valAx>');
    expect(ticks(x!)).toEqual(['out', 'none']);
    expect(ticks(y!)).toEqual(['out', 'none']);
  });

  it('writes outward category ticks on a radar chart', async () => {
    const xml = await chartXml({ kind: 'radar', categories: ['A', 'B', 'C'], series });
    expect(ticks(axis(xml, 'catAx'))).toEqual(['out', 'none']);
  });

  it('decides the category ticks by the base kind of a combo chart', async () => {
    const xml = await chartXml({
      kind: 'line',
      categories: ['A', 'B'],
      series: [...series, { name: 'T', values: [3, 4], chartKind: 'column' }],
    });
    expect(ticks(axis(xml, 'catAx'))).toEqual(['out', 'none']);
  });

  it('writes outward ticks on the series axis of a surface chart', async () => {
    const xml = await chartXml({
      kind: 'surface',
      categories: ['A', 'B'],
      series: [...series, { name: 'T', values: [3, 4] }],
    });
    expect(ticks(axis(xml, 'serAx'))).toEqual(['out', 'none']);
  });

  it('keeps authored tick marks', async () => {
    const xml = await chartXml({
      kind: 'column',
      categories: ['A', 'B'],
      series,
      categoryAxisMajorTickMark: 'cross',
      categoryAxisMinorTickMark: 'in',
      valueAxisMajorTickMark: 'none',
      valueAxisMinorTickMark: 'out',
    });
    expect(ticks(axis(xml, 'catAx'))).toEqual(['cross', 'in']);
    expect(ticks(axis(xml, 'valAx'))).toEqual(['none', 'out']);
  });
});
