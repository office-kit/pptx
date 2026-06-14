// Tests for two correctness fixes:
//
//   - Unit conversions and shape geometry always emit whole EMU. Fractional
//     ST_Coordinate values (floating-point drift from unit math, or an
//     `as Emu` cast on a computed value) are schema-invalid and make
//     PowerPoint mark the file corrupt and zero the offending offsets.
//   - Stacked / 100%-stacked bar charts default to overlap=100 so the stack
//     doesn't spread sideways across the category.

import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideChart,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  cm,
  createPresentation,
  emu,
  findSlideLayout,
  getSlides,
  getSlideXmlString,
  inches,
  mm,
  pt,
} from '../src/api/index.ts';

const blankSlide = () => {
  const pres = createPresentation();
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected Blank layout');
  return { pres, slide: addSlide(pres, { layout }) };
};

// Coordinate attributes (x / y / cx / cy) that carry a fractional value.
const fractionalCoords = (xml: string): string[] =>
  [...xml.matchAll(/[^a-z](x|y|cx|cy)="(-?\d+\.\d+)"/g)].map((m) => `${m[1]}=${m[2]}`);

const chartXml = (pres: ReturnType<typeof createPresentation>, n: number): string => {
  const part = _internalPackageOf(pres).parts.find((p) => p.name === `/ppt/charts/chart${n}.xml`);
  if (!part) throw new Error(`chart${n} part not found`);
  return new TextDecoder().decode(part.data);
};

describe('units: whole-EMU conversion', () => {
  it('inches rounds floating-point drift to an integer', () => {
    // 3.379 * 914400 = 3090677.6 in exact math but FP gives ...0000005-style drift.
    const v = inches(3.379);
    expect(Number.isInteger(v)).toBe(true);
    expect(v).toBe(Math.round(3.379 * 914400));
  });

  it('cm / mm / pt return integers', () => {
    expect(Number.isInteger(cm(1.234))).toBe(true);
    expect(Number.isInteger(mm(12.34))).toBe(true);
    expect(Number.isInteger(pt(0.75))).toBe(true);
    expect(pt(0.75)).toBe(9525);
  });

  it('emu escape hatch rounds a fractional value', () => {
    expect(emu(1000.5)).toBe(1001);
    expect(Number.isInteger(emu(123.4))).toBe(true);
  });

  it('clean values are unchanged', () => {
    expect(inches(1)).toBe(914400);
    expect(inches(0.5)).toBe(457200);
  });
});

describe('shape geometry: no fractional EMU reaches the XML', () => {
  it('text box / autoshape / table offsets are whole EMU', () => {
    const { pres, slide } = blankSlide();
    addSlideShape(slide, {
      preset: 'rect',
      x: inches(3.379),
      y: inches(3.379),
      w: inches(4.455),
      h: inches(3.64),
      text: 'X',
    });
    // emu() escape hatch with fractional inputs must still serialize as integers.
    addSlideTextBox(slide, {
      x: emu(123.7),
      y: emu(456.4),
      w: emu(1000.5),
      h: emu(2000.5),
      text: 'Y',
    });
    addSlideTable(slide, {
      x: inches(0.31),
      y: inches(0.69),
      w: inches(5.13),
      h: inches(1.27),
      rows: [['a', 'b']],
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(fractionalCoords(xml)).toEqual([]);
    // Spot-check the rounded autoshape origin.
    expect(xml).toContain(`x="${Math.round(inches(3.379))}"`);
  });

  it('chart graphic frame offset is whole EMU', () => {
    const { pres, slide } = blankSlide();
    addSlideChart(slide, {
      x: emu(457200.6),
      y: emu(457200.4),
      w: inches(6),
      h: inches(4),
      spec: { kind: 'column', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] },
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(fractionalCoords(xml)).toEqual([]);
  });
});

describe('bar chart overlap defaults', () => {
  it('stacked column defaults to overlap=100', () => {
    const { pres, slide } = blankSlide();
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        grouping: 'stacked',
        categories: ['a', 'b'],
        series: [
          { name: 's1', values: [1, 2] },
          { name: 's2', values: [3, 4] },
        ],
      },
    });
    expect(chartXml(pres, 1)).toContain('<c:overlap val="100"/>');
  });

  it('percentStacked column defaults to overlap=100', () => {
    const { pres, slide } = blankSlide();
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        grouping: 'percentStacked',
        categories: ['a', 'b'],
        series: [
          { name: 's1', values: [1, 2] },
          { name: 's2', values: [3, 4] },
        ],
      },
    });
    expect(chartXml(pres, 1)).toContain('<c:overlap val="100"/>');
  });

  it('an explicit overlapPct wins over the stacked default', () => {
    const { pres, slide } = blankSlide();
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        grouping: 'stacked',
        overlapPct: 50,
        categories: ['a', 'b'],
        series: [{ name: 's1', values: [1, 2] }],
      },
    });
    const xml = chartXml(pres, 1);
    expect(xml).toContain('<c:overlap val="50"/>');
    expect(xml).not.toContain('<c:overlap val="100"/>');
  });

  it('clustered column emits no default overlap (keeps PowerPoint default)', () => {
    const { pres, slide } = blankSlide();
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: { kind: 'column', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }] },
    });
    expect(chartXml(pres, 1)).not.toContain('<c:overlap');
  });
});
