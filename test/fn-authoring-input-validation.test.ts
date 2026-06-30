// Authoring-boundary input validation for charts, tables, and layout lookup.
//
// Regression coverage for the issues surfaced integrating pptx-kit into a
// downstream app:
//   - chart series colors that aren't sRGB hex were silently emitted as
//     invalid `<a:srgbClr val="…"/>` (PowerPoint dropped/repaired them);
//   - `addSlideTable` with empty `rows: []` produced a table with no grid;
//   - `findSlideLayout` name matching is case-sensitive (documented, not
//     changed — this test pins the behavior so it doesn't drift).

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addSlideChart,
  addSlideTable,
  createPresentation,
  findSlideLayout,
  findSlideLayoutByType,
  getSlideShapes,
  inches,
  type ChartSpec,
} from '../src/api/index.ts';

const columnSpec = (color: string): ChartSpec => ({
  kind: 'column',
  categories: ['A', 'B'],
  series: [{ name: 'S', values: [1, 2], color }],
});

describe('addSlideChart: series color validation', () => {
  it('rejects a non-hex color with a clear, actionable error', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: columnSpec('red'),
      }),
    ).toThrow(/series\[0\]\.color: invalid chart color "red" — expected an sRGB hex/);
  });

  it('rejects a malformed hex (wrong length / non-hex digits)', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    for (const bad of ['#GGGGGG', 'rgb(0,0,0)', '#12345', '#FFFF']) {
      expect(() =>
        addSlideChart(slide, {
          x: inches(1),
          y: inches(1),
          w: inches(4),
          h: inches(3),
          spec: columnSpec(bad),
        }),
      ).toThrow(/invalid chart color/);
    }
  });

  it('rejects scheme tokens — charts emit srgbClr, not schemeClr', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: columnSpec('accent1'),
      }),
    ).toThrow(/invalid chart color/);
  });

  it('accepts both #RRGGBB and bare RRGGBB', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: columnSpec('#4472C4'),
      }),
    ).not.toThrow();
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: columnSpec('4472C4'),
      }),
    ).not.toThrow();
  });

  it('accepts the 3-digit hex shorthand', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    // `#4cf` is the CSS-style shorthand LLM authors reach for; it must be
    // accepted (and expanded to `44CCFF` internally), not rejected.
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: columnSpec('#4cf'),
      }),
    ).not.toThrow();
  });

  it('validates pointColors and trendline.color too', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideChart(slide, {
        x: inches(1),
        y: inches(1),
        w: inches(4),
        h: inches(3),
        spec: {
          kind: 'pie',
          categories: ['A', 'B'],
          series: [{ name: 'S', values: [1, 2], pointColors: ['#FF0000', 'notacolor'] }],
        },
      }),
    ).toThrow(/pointColors\[1\]/);
  });
});

describe('addSlideTable: empty rows validation', () => {
  it('throws an actionable error on empty rows instead of emitting broken XML', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const before = getSlideShapes(slide).length;
    expect(() =>
      addSlideTable(slide, { x: inches(1), y: inches(1), w: inches(4), h: inches(2), rows: [] }),
    ).toThrow(/addSlideTable: at least one row is required/);
    // The slide must be untouched — no half-built table shape appended.
    expect(getSlideShapes(slide).length).toBe(before);
  });

  it('throws on a row with no cells', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    expect(() =>
      addSlideTable(slide, { x: inches(1), y: inches(1), w: inches(4), h: inches(2), rows: [[]] }),
    ).toThrow(/addSlideTable: at least one column is required/);
  });

  it('still builds a normal table when rows are non-empty', () => {
    const pres = createPresentation();
    const slide = addBlankSlide(pres);
    const before = getSlideShapes(slide).length;
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [
        ['H1', 'H2'],
        ['a', 'b'],
      ],
    });
    expect(getSlideShapes(slide).length).toBe(before + 1);
  });
});

describe('findSlideLayout: documented case sensitivity', () => {
  it('matches the exact, case-sensitive user-visible name', () => {
    const pres = createPresentation();
    expect(findSlideLayout(pres, 'Blank')).not.toBeNull();
    // Case mismatch does NOT match (this is the documented behavior).
    expect(findSlideLayout(pres, 'blank')).toBeNull();
  });

  it('a case-insensitive RegExp bridges the gap', () => {
    const pres = createPresentation();
    expect(findSlideLayout(pres, /^blank$/i)).not.toBeNull();
  });

  it('findSlideLayoutByType matches the locale-stable token regardless of name case', () => {
    const pres = createPresentation();
    expect(findSlideLayoutByType(pres, 'blank')).not.toBeNull();
  });
});
