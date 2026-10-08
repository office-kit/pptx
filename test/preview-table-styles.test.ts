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
  setPresentationTheme,
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
const channelDistance = (a: string, b: string): number =>
  Math.max(
    ...[0, 2, 4].map((i) =>
      Math.abs(Number.parseInt(a.slice(i, i + 2), 16) - Number.parseInt(b.slice(i, i + 2), 16)),
    ),
  );
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

  // Medium Style 2's bands are accent1 tinted 40% / 20%. In the Office 2023
  // theme (accent1 #156082) PowerPoint paints them #CCD2D8 / #E6EAED (native
  // Table Design screenshot, Display P3 converted to sRGB); an sRGB tint
  // would give the much bluer #A1BFCD / #D0DFE6. One level of tolerance for
  // the color-space conversion.
  it('paints Medium Style 2’s tinted bands as PowerPoint does in the Office 2023 theme', () => {
    const { pres, svg } = render('Medium Style 2 - Accent 1', { firstRow: true, bandRow: true });
    setPresentationTheme(pres, { accent1: '#156082' });
    const expected = ['156082', 'CCD2D8', 'E6EAED', 'CCD2D8', 'E6EAED'];
    const actual = column(svg(), 0);
    expected.forEach((color, r) =>
      expect(channelDistance(actual[r]!, color), `row ${r}: ${actual[r]}`).toBeLessThanOrEqual(1),
    );
  });

  // The Themed Styles' backgrounds are the theme's gradient fills, whose
  // stops saturate accents with satMod 130–350%. Expected stop colors come
  // from Mac PowerPoint 16.113's exports (Office 2007 theme), sampled 3 px
  // from each end of the table; 2 levels of tolerance for that offset.
  it.each([
    ['Themed Style 1 - Accent 6', 'FFBE87', undefined],
    ['Themed Style 1 - Accent 1', 'A3C4FF', undefined],
    ['Themed Style 2 - Accent 6', 'FF932C', 'FFB977'],
    ['Themed Style 2 - Accent 5', '3AB8D8', '95EEFF'],
  ] as const)('saturates %s’s background gradient like PowerPoint', (style, first, last) => {
    const { svg } = render(style, { firstRow: true, bandRow: true });
    const out = svg();
    const gradient = out.slice(out.indexOf('<linearGradient'), out.indexOf('</linearGradient>'));
    const stops = [...gradient.matchAll(/stop-color="#([0-9A-F]{6})"/gi)].map((m) => m[1]!);
    expect(channelDistance(stops[0]!, first), stops[0]).toBeLessThanOrEqual(2);
    if (last !== undefined)
      expect(channelDistance(stops.at(-1)!, last), stops.at(-1)).toBeLessThanOrEqual(2);
  });

  // Themed Style 2's two-stop background does not blend at a constant rate in
  // PowerPoint. Reference pixels: Mac PowerPoint 16.113's 1200 × 700 exports,
  // column x = 1000 of a table spanning y = 100–500, in the header and the two
  // unbanded body rows; keys are the pixel centres as a fraction of the table
  // height from its top. 2 levels of tolerance: the fitted curve's worst
  // residual over all six accents is 1.6.
  it.each([
    [
      'Themed Style 2 - Accent 1',
      [
        [0.0512, '9BC1FF'],
        [0.1512, '98BFFE'],
        [0.4512, '86B2F4'],
        [0.5513, '7DACEF'],
        [0.8512, '5791DA'],
        [0.9513, '4786D1'],
      ],
    ],
    [
      'Themed Style 2 - Accent 6',
      [
        [0.0512, 'FFB977'],
        [0.1512, 'FFB875'],
        [0.4512, 'FFB066'],
        [0.5513, 'FFAC5E'],
        [0.8512, 'FF9D3F'],
        [0.9513, 'FF9632'],
      ],
    ],
  ] as const)('blends %s’s background down the table as PowerPoint does', (style, samples) => {
    const { svg } = render(style, { firstRow: true, bandRow: true });
    const out = svg();
    const gradient = out.slice(out.indexOf('<linearGradient'), out.indexOf('</linearGradient>'));
    const attr = (name: string) => Number(new RegExp(` ${name}="([-\\d.]+)"`).exec(gradient)![1]);
    // A vertical gradient in the table's bounding box.
    expect(attr('x1')).toBe(attr('x2'));
    const [y1, y2] = [attr('y1'), attr('y2')];
    const stops = [...gradient.matchAll(/offset="([\d.]+)" stop-color="#([0-9A-F]{6})"/gi)].map(
      (m) => ({
        offset: Number(m[1]),
        rgb: [0, 2, 4].map((i) => parseInt(m[2]!.slice(i, i + 2), 16)),
      }),
    );
    // SVG interpolates between stops in sRGB, as PowerPoint does here.
    const colorAt = (fraction: number): string => {
      const offset = (fraction - y1) / (y2 - y1);
      const after = stops.findIndex((stop) => stop.offset >= offset);
      const b = stops[after]!;
      const a = stops[Math.max(0, after - 1)]!;
      const t = b.offset === a.offset ? 0 : (offset - a.offset) / (b.offset - a.offset);
      return a.rgb
        .map((v, k) => Math.round(v + (b.rgb[k]! - v) * t))
        .map((v) => v.toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
    };
    for (const [fraction, expected] of samples) {
      const actual = colorAt(fraction);
      expect(channelDistance(actual, expected), `${fraction}: ${actual}`).toBeLessThanOrEqual(2);
    }
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
