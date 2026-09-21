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

const CATEGORY_COLOR = '#11AA11';
const VALUE_COLOR = '#AA1111';
const frame = { left: 136, top: 104, width: 720, height: 446, bottom: 550 };
async function render(spec: ChartSpec) {
  const pres = await loadPresentation(
    await readFile(new URL('./fixtures/minimal/blank.pptx', import.meta.url)),
  );
  const slide = addSlide(pres, { layout: findSlideLayout(pres, 'Blank')! });
  addSlideChart(slide, { x: inches(1), y: inches(1), w: inches(8), h: inches(5), spec });
  return renderSlideToSvg(pres, slide);
}
function chart(kind: ChartSpec['kind'] | 'combo'): ChartSpec {
  return {
    kind: kind === 'combo' ? 'column' : kind,
    categories: ['Alpha', 'Beta'],
    valueAxis: { min: 0, max: 100, majorUnit: 20 },
    valueAxisLineColor: VALUE_COLOR,
    categoryAxisLineColor: CATEGORY_COLOR,
    series: [
      { name: 'S', values: [20, 60], xValues: [1, 2], bubbleSizes: [5, 10] },
      ...(kind === 'combo' ? [{ name: 'Line', values: [30, 40], chartKind: 'line' as const }] : []),
    ],
  };
}
const spines = (svg: string) =>
  attrsOf(svg, 'line').filter(
    (l) => Math.abs(Number(l.x1) - Number(l.x2)) > 20 || Math.abs(Number(l.y1) - Number(l.y2)) > 20,
  );

describe.each(['column', 'bar', 'line', 'area', 'combo'] as const)('%s axis visibility', (kind) => {
  it('hides only the spines, retaining labels and ticks', async () => {
    const svg = await render({
      ...chart(kind),
      valueAxisLineHidden: true,
      categoryAxisLineHidden: true,
    });
    expect(spines(svg)).toHaveLength(0);
    expect(svg).toContain('>Alpha</text>');
    expect(svg).toContain('>100</text>');
    expect(attrsOf(svg, 'line')).toHaveLength(9);
  });
  it.each(['value', 'category'] as const)('hides only the %s axis spine', async (axis) => {
    const svg = await render({
      ...chart(kind),
      valueAxisLineHidden: axis === 'value',
      categoryAxisLineHidden: axis === 'category',
    });
    const coloredSpines = spines(svg).filter(
      (l) => l.stroke === CATEGORY_COLOR || l.stroke === VALUE_COLOR,
    );
    expect(coloredSpines).toHaveLength(1);
    expect(coloredSpines[0]?.stroke).toBe(axis === 'value' ? CATEGORY_COLOR : VALUE_COLOR);
    expect(svg).toContain('>Alpha</text>');
    expect(svg).toContain('>100</text>');
  });
  it('keeps explicitly visible spines', async () => {
    const svg = await render({
      ...chart(kind),
      valueAxisLineHidden: false,
      categoryAxisLineHidden: false,
    });
    expect(
      spines(svg).filter((l) => l.stroke === CATEGORY_COLOR || l.stroke === VALUE_COLOR),
    ).toHaveLength(2);
  });
  it('still removes the entire axis when AxisHidden is set', async () => {
    const svg = await render({ ...chart(kind), valueAxisHidden: true, categoryAxisHidden: true });
    expect(spines(svg)).toHaveLength(0);
    expect(svg).not.toContain('>Alpha</text>');
    expect(svg).not.toContain('>100</text>');
  });
});

describe.each(['column', 'bar', 'combo'] as const)('%s category ticks', (kind) => {
  describe.each(['major', 'minor'] as const)('%s marks', (level) => {
    it.each(['none', 'in', 'out', 'cross'] as const)(
      'renders %s marks at the correct positions',
      async (mark) => {
        const svg = await render({
          ...chart(kind),
          valueAxisHidden: true,
          categoryAxisLineHidden: true,
          categoryAxisMajorTickMark: level === 'major' ? mark : 'none',
          categoryAxisMinorTickMark: level === 'minor' ? mark : 'none',
        });
        const ticks = attrsOf(svg, 'line');
        const count = level === 'major' ? 3 : 2;
        expect(ticks).toHaveLength(mark === 'none' ? 0 : count);
        const length = level === 'major' ? 5 : 3;
        const outward = mark === 'in' ? 0 : length;
        const inward = mark === 'out' ? 0 : -length;
        ticks.forEach((tick, i) => {
          const fraction = (i + (level === 'minor' ? 0.5 : 0)) / 2;
          if (kind === 'bar') {
            expect(Number(tick.x1)).toBeCloseTo(frame.left - outward);
            expect(Number(tick.x2)).toBeCloseTo(frame.left - inward);
            expect(Number(tick.y1)).toBeCloseTo(frame.top + frame.height * fraction);
            expect(tick.y2).toBe(tick.y1);
          } else {
            expect(Number(tick.y1)).toBeCloseTo(frame.bottom + outward);
            expect(Number(tick.y2)).toBeCloseTo(frame.bottom + inward);
            expect(Number(tick.x1)).toBeCloseTo(frame.left + frame.width * fraction);
            expect(tick.x2).toBe(tick.x1);
          }
        });
        expect(svg).toContain('>Alpha</text>');
      },
    );
  });
  it('keeps the default outward major ticks and no minor ticks', async () => {
    const svg = await render({
      ...chart(kind),
      valueAxisHidden: true,
      categoryAxisLineHidden: true,
    });
    const ticks = attrsOf(svg, 'line');
    expect(ticks).toHaveLength(3);
    expect(
      ticks.every(
        (l) => Math.abs(Number(l.x1) - Number(l.x2)) + Math.abs(Number(l.y1) - Number(l.y2)) === 5,
      ),
    ).toBe(true);
  });
});

describe.each(['scatter', 'bubble'] as const)('%s numeric axis spines', (kind) => {
  it('honors line-only visibility independently of whole-axis visibility', async () => {
    const svg = await render({
      ...chart(kind),
      valueAxisLineHidden: true,
      categoryAxisLineHidden: true,
    });
    expect(spines(svg)).toHaveLength(0);
    expect(svg).toContain('>100</text>');
    expect(attrsOf(svg, 'line').length).toBeGreaterThan(0);
    const hidden = await render({
      ...chart(kind),
      valueAxisHidden: true,
      categoryAxisHidden: true,
    });
    expect(attrsOf(hidden, 'line')).toHaveLength(0);
    expect(hidden).not.toContain('>100</text>');
  });
});

it.each([0.01, 1e-100, Number.MIN_VALUE])(
  'keeps dense value-axis interval %s responsive',
  async (majorUnit) => {
    const svg = await render({
      ...chart('column'),
      valueAxis: { min: 0, max: 100, majorUnit },
      valueAxisMajorGridlines: true,
    });
    expect(attrsOf(svg, 'line').length).toBeLessThan(2100);
    expect(svg).toContain('>0</text>');
    expect(svg).toContain('>100</text>');
  },
);

it('renders fractional authored ticks without accumulation drift', async () => {
  const svg = await render({
    ...chart('bar'),
    valueAxis: { min: -0.3, max: 0.3, majorUnit: 0.1, numberFormat: '0.0' },
    valueAxisMajorGridlines: true,
  });
  const grid = attrsOf(svg, 'line').filter((line) => line['stroke-width'] === '0.5');
  expect(grid).toHaveLength(7);
  expect(svg).toContain('>0.0</text>');
});

it('finishes automatic ticks when increments are below floating-point precision', async () => {
  const svg = await render({
    ...chart('column'),
    valueAxis: { min: 1e16, max: 1e16 + 2 },
    series: [{ name: 'Large values', values: [1e16, 1e16 + 2] }],
  });
  expect(attrsOf(svg, 'line').length).toBeLessThan(2100);
  expect(svg).not.toMatch(/(?:NaN|Infinity)/);
});

describe.each(['column', 'bar', 'line', 'area', 'combo'] as const)('%s minor gridlines', (kind) => {
  it('renders authored spacing, color and width without duplicating major gridlines', async () => {
    const svg = await render({
      ...chart(kind),
      valueAxis: { min: 0, max: 100, majorUnit: 20, minorUnit: 4 },
      valueAxisMajorGridlines: true,
      valueAxisMinorGridlines: true,
      valueAxisMinorGridlineColor: '#123456',
      valueAxisMinorGridlineWidthEmu: 9525,
    });
    const minor = attrsOf(svg, 'line').filter((line) => line.stroke === '#123456');
    expect(minor).toHaveLength(20);
    minor.forEach((line, i) => {
      const fraction = (Math.floor(i / 4) * 20 + ((i % 4) + 1) * 4) / 100;
      expect(Number(line['stroke-width'])).toBe(1);
      if (kind === 'bar') {
        expect(line.x1).toBe(line.x2);
        expect(Number(line.x1)).toBeCloseTo(frame.left + frame.width * fraction);
      } else {
        expect(line.y1).toBe(line.y2);
        expect(Number(line.y1)).toBeCloseTo(frame.bottom - frame.height * fraction);
      }
    });
    expect(attrsOf(svg, 'line').filter((line) => line.stroke === '#D9D9D9')).toHaveLength(6);
  });
});

it.each([undefined, 1e-100])(
  'uses bounded automatic minor spacing for interval %s',
  async (minorUnit) => {
    const svg = await render({
      ...chart('column'),
      valueAxis: {
        min: 0,
        max: 100,
        majorUnit: 20,
        ...(minorUnit === undefined ? {} : { minorUnit }),
      },
      valueAxisMinorGridlines: true,
    });
    expect(
      attrsOf(svg, 'line').filter((line) => line['data-chart-gridline'] === 'minor'),
    ).toHaveLength(20);
  },
);

it.each([false, true])(
  'does not draw minor gridlines when disabled or the axis is hidden (%s)',
  async (hidden) => {
    const svg = await render({
      ...chart('column'),
      valueAxisMinorGridlines: hidden,
      valueAxisHidden: hidden,
    });
    expect(svg).not.toContain('data-chart-gridline="minor"');
  },
);
