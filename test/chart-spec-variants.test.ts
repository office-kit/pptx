// `ChartSpec` is a discriminated union: each kind carries only the fields its
// OOXML element has. These tests pin the two ways that can go wrong — a field
// no variant lists (unusable), and a guard that disagrees with the type.
import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { type ChartSpec, type ReadChartSpec, isChartSpec } from '../src/api/index.ts';
import { ALL_VARIANT_FIELDS } from '../src/internal/chartml/index.ts';

const TYPES_SOURCE = new URL('../src/internal/chartml/types.ts', import.meta.url);

// Top-level members of `ReadChartSpec`; nested object fields are indented further.
const readSpecFields = async (): Promise<ReadonlyArray<string>> => {
  const source = await readFile(TYPES_SOURCE, 'utf8');
  const start = source.indexOf('export interface ReadChartSpec {');
  expect(start).toBeGreaterThan(-1);
  const body = source.slice(start, source.indexOf('\n}\n', start));
  return [...body.matchAll(/^ {2}readonly ([A-Za-z0-9]+)\??:/gm)].map((m) => m[1]!);
};

describe('ChartSpec variant field coverage', () => {
  it('leaves no field of ReadChartSpec that no kind can set', async () => {
    const unusable = (await readSpecFields()).filter(
      (field) => field !== 'kind' && field !== 'series' && !ALL_VARIANT_FIELDS.has(field),
    );
    expect(unusable).toEqual([]);
  });

  it('lists no field that ReadChartSpec does not declare', async () => {
    const declared = new Set(await readSpecFields());
    expect([...ALL_VARIANT_FIELDS].filter((field) => !declared.has(field))).toEqual([]);
  });
});

const PIE: ChartSpec = { kind: 'pie', categories: ['A'], series: [{ name: 'S', values: [1] }] };
const COLUMN: ChartSpec = {
  kind: 'column',
  categories: ['A'],
  series: [{ name: 'S', values: [1] }],
};

describe('isChartSpec', () => {
  it('accepts what the write side can author', () => {
    expect(isChartSpec(PIE)).toBe(true);
    expect(isChartSpec(COLUMN)).toBe(true);
  });

  // Each case is a spec a deck could carry but no kind draws — the same
  // combinations `ChartSpec` rejects at compile time.
  it.each([
    ['an axis field on a pie', { ...PIE, valueAxisTitle: 'Revenue' }],
    ['a scatter subtype on a column', { ...COLUMN, scatterStyle: 'line' as const }],
    ['bar3DShape without view3D', { ...COLUMN, bar3DShape: 'cone' as const }],
    ['a 3-D pie with firstSliceAngleDeg', { ...PIE, view3D: {}, firstSliceAngleDeg: 90 }],
    ['ofPie on a doughnut', { ...PIE, kind: 'doughnut' as const, ofPie: { type: 'pie' as const } }],
    ['a pie with two series', { ...PIE, series: [...PIE.series, { name: 'T', values: [2] }] }],
    ['a scatter series with no x channel', { ...COLUMN, kind: 'scatter' as const }],
    [
      'a date axis with outer category levels',
      { ...COLUMN, categoryAxisDate: {}, categoryGroupLevels: [['Y']] },
    ],
    ['a text axis with a scale', { ...COLUMN, categoryAxisScaling: { min: 0 } }],
  ])('rejects %s', (_label, spec) => {
    expect(isChartSpec(spec satisfies ReadChartSpec)).toBe(false);
  });

  it('keeps a numeric horizontal axis scalable', () => {
    const scatter: ReadChartSpec = {
      kind: 'scatter',
      categories: [],
      series: [{ name: 'S', values: [1], xValues: [1] }],
      categoryAxisScaling: { min: 0, max: 10 },
    };
    expect(isChartSpec(scatter)).toBe(true);
  });

  it('tells a stock chart from a candlestick by its series count', () => {
    const series = Array.from({ length: 3 }, (_, i) => ({ name: `S${i}`, values: [1] }));
    const stock: ReadChartSpec = { kind: 'stock', categories: ['Mon'], series };
    expect(isChartSpec(stock)).toBe(true);
    // Up / down bars span open → close, so they need the fourth series.
    expect(isChartSpec({ ...stock, upDownBars: {} })).toBe(false);
    expect(
      isChartSpec({ ...stock, series: [...series, { name: 'S3', values: [1] }], upDownBars: {} }),
    ).toBe(true);
  });
});

describe('ChartSpec compile-time rejections', () => {
  it('will not let a kind carry a field it cannot draw', () => {
    // @ts-expect-error a pie has no axes
    const axisOnPie: ChartSpec = { ...PIE, valueAxisTitle: 'Revenue' };
    // @ts-expect-error <c:shape> lives on <c:bar3DChart>, which needs view3D
    const shapeWithout3D: ChartSpec = { ...COLUMN, bar3DShape: 'cone' };
    // @ts-expect-error a combo series cannot share a plot area with a 3-D group
    const comboIn3D: ChartSpec = {
      ...COLUMN,
      view3D: {},
      series: [{ name: 'S', values: [1], chartKind: 'line' }],
    };
    // @ts-expect-error a surface chart emits no <c:gapDepth>
    const gapDepthOnSurface: ChartSpec = {
      ...COLUMN,
      kind: 'surface',
      view3D: {},
      gapDepthPct: 50,
    };
    expect([axisOnPie, shapeWithout3D, comboIn3D, gapDepthOnSurface]).toHaveLength(4);
  });
});
