// Preview rendering of PowerPoint's built-in table styles: cell fills, the
// table background, borders and text across the six style options.
//
// The expected colors were sampled from Mac PowerPoint 16.113's exports of
// tables in the Office 2007 theme the blank deck uses (header row + banded
// rows on), so they are PowerPoint's colors rather than restated formulas.

import { describe, expect, it } from 'vitest';
import {
  type BuiltinTableStyleName,
  addBlankSlide,
  addSlideTable,
  createPresentation,
  getTableCellAppearanceEffective,
  getTableCellRunFormatEffective,
  getTableCells,
  inches,
  setTableCellFill,
  setTableCellTextFormat,
  setTableStyleFlags,
  setTableStyleId,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

type Flags = Parameters<typeof setTableStyleFlags>[1];

const render = (style: BuiltinTableStyleName, flags: Flags = {}) => {
  const pres = createPresentation();
  const slide = addBlankSlide(pres);
  const table = addSlideTable(slide, {
    x: inches(0.5),
    y: inches(1),
    w: inches(8),
    h: inches(4),
    rows: Array.from({ length: 5 }, (_, r) => Array.from({ length: 4 }, (_, c) => `${r}.${c}`)),
  });
  setTableStyleId(table, style);
  setTableStyleFlags(table, flags);
  return { pres, slide, table, svg: () => renderSlideToSvg(pres, slide) };
};

// Each cell's painted color, composited over the white slide.
const cellColors = (svg: string): string[][] => {
  const out: string[][] = [];
  for (const m of svg.matchAll(
    /<g data-pptx-cell="(\d+),(\d+)"><rect [^>]*fill="([^"]+)"(?: fill-opacity="([\d.]+)")?/g,
  )) {
    const [, r, c, fill, opacity] = m;
    const alpha = opacity === undefined ? 1 : Number(opacity);
    const color =
      fill === 'none'
        ? 'FFFFFF'
        : [1, 3, 5]
            .map((i) => Number.parseInt(fill!.slice(i, i + 2), 16))
            .map((v) => Math.round(v * alpha + 255 * (1 - alpha)))
            .map((v) => v.toString(16).padStart(2, '0'))
            .join('')
            .toUpperCase();
    (out[Number(r)] ??= [])[Number(c)] = color;
  }
  return out;
};
const column = (svg: string, col: number) => cellColors(svg).map((row) => row[col]);
const row = (svg: string, r: number) => cellColors(svg)[r];

describe('built-in table styles in the preview', () => {
  it.each([
    ['Medium Style 2 - Accent 1', ['4F81BD', 'D0D8E8', 'E9EDF4', 'D0D8E8', 'E9EDF4']],
    ['Light Style 1 - Accent 1', ['FFFFFF', 'DCE6F2', 'FFFFFF', 'DCE6F2', 'FFFFFF']],
    ['Light Style 2 - Accent 1', ['4F81BD', 'FFFFFF', 'FFFFFF', 'FFFFFF', 'FFFFFF']],
    ['Medium Style 3 - Accent 1', ['4F81BD', 'E7E7E7', 'FFFFFF', 'E7E7E7', 'FFFFFF']],
    ['Medium Style 4 - Accent 1', ['E9EDF4', 'D0D8E8', 'E9EDF4', 'D0D8E8', 'E9EDF4']],
    ['Dark Style 1', ['000000', 'CBCBCB', 'E7E7E7', 'CBCBCB', 'E7E7E7']],
    ['Dark Style 2 - Accent 1/Accent 2', ['C0504D', 'D0D8E8', 'E9EDF4', 'D0D8E8', 'E9EDF4']],
  ] as const)('%s paints PowerPoint’s header and banded rows', (style, expected) => {
    const { svg } = render(style, { firstRow: true, bandRow: true });
    for (let c = 0; c < 4; c++) expect(column(svg(), c)).toEqual(expected);
  });

  it('applies first/last column, total row and banded columns (Medium Style 2)', () => {
    const { svg } = render('Medium Style 2 - Accent 1', {
      firstRow: true,
      lastRow: true,
      firstCol: true,
      lastCol: true,
      bandRow: false,
      bandCol: true,
    });
    const accent = '4F81BD';
    // Header and total rows win over the columns in the corners.
    expect(row(svg(), 0)).toEqual([accent, accent, accent, accent]);
    expect(row(svg(), 4)).toEqual([accent, accent, accent, accent]);
    // Body: first and last columns are solid, the two between band from the
    // first one after the first column.
    for (const r of [1, 2, 3]) expect(row(svg(), r)).toEqual([accent, 'D0D8E8', 'E9EDF4', accent]);
  });

  it('bands rows from the first body row and stops before the total row', () => {
    const { svg } = render('Medium Style 2 - Accent 1', {
      firstRow: false,
      lastRow: true,
      bandRow: true,
    });
    expect(column(svg(), 0)).toEqual(['D0D8E8', 'E9EDF4', 'D0D8E8', 'E9EDF4', '4F81BD']);
  });

  it('draws nothing but the style’s whole-table fill with every option off', () => {
    const { svg } = render('Medium Style 2 - Accent 1', { firstRow: false, bandRow: false });
    expect(new Set(cellColors(svg()).flat())).toEqual(new Set(['E9EDF4']));
  });

  it('lets direct cell formatting win over the style', () => {
    const { table, svg } = render('Medium Style 2 - Accent 1', { firstRow: true, bandRow: true });
    setTableCellFill(getTableCells(table)[0]![1]!, '#FF0000');
    setTableCellFill(getTableCells(table)[2]![1]!, '#00FF00');
    expect(row(svg(), 0)![1]).toBe('FF0000');
    expect(row(svg(), 2)![1]).toBe('00FF00');
  });

  it('resolves header borders and text from the style parts', () => {
    const { pres, table } = render('Medium Style 2 - Accent 1', { firstRow: true, bandRow: true });
    const header = getTableCells(table)[0]![0]!;
    expect(getTableCellAppearanceEffective(pres, header).borders.bottom).toEqual({
      color: '#FFFFFF',
      widthEmu: 38100,
      dash: null,
    });
    // addSlideTable bakes the deck's text color into runs; clear it so the
    // style's text color shows.
    setTableCellTextFormat(header, {}, { reset: true });
    const format = getTableCellRunFormatEffective(pres, header, 0, 0);
    expect(format.bold).toBe(true);
    expect(format.color).toBe('#FFFFFF');
    const body = getTableCells(table)[1]![0]!;
    setTableCellTextFormat(body, {}, { reset: true });
    expect(getTableCellRunFormatEffective(pres, body, 0, 0).bold).toBeUndefined();
  });

  it('paints the Themed Styles’ theme gradient behind translucent bands', () => {
    const { svg } = render('Themed Style 1 - Accent 1', { firstRow: true, bandRow: true });
    const out = svg();
    const firstCell = out.indexOf('data-pptx-cell="0,0"');
    const gradient = out.indexOf('<linearGradient');
    expect(gradient).toBeGreaterThan(-1);
    expect(gradient).toBeLessThan(firstCell);
    // band1H: accent1 at 40% alpha over the background; band2H: no fill.
    expect(out).toMatch(/data-pptx-cell="1,0"><rect [^>]*fill="#4F81BD" fill-opacity="0.400"/);
    expect(out).toMatch(/data-pptx-cell="2,0"><rect [^>]*fill="none"/);
  });
});
