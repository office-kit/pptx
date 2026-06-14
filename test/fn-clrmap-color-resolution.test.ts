// Tests for clrMap-aware color resolution and the baking of body-text color
// onto authored tables and charts.
//
// Covered behaviors (added in fix/clrmap-aware-color-resolution):
//
//   a. resolveDrawingColor + clrMap parameter — schemeClr tokens are remapped
//      through the clrMap before the theme lookup.
//   b. getEffectiveColorMap — returns the standard map for a createPresentation
//      deck (master clrMap bg1="lt1" tx1="dk1").
//   c. addSlideTable bakes color — cell runs carry the resolved body-text color
//      as an explicit <a:solidFill> when the deck has a theme.
//   d. addSlideChart bakes color — catAx / valAx (and legend / dLbls when
//      present) carry the resolved body-text color when no authored color is set.
//   e. Inverted clrMap end-to-end — rewriting the master's <p:clrMap> to an
//      inverted form makes getEffectiveColorMap report the swap, and the baked
//      table / chart colors follow (tx1 → lt1 → FFFFFF on the standard theme).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideChart,
  addSlideTable,
  createPresentation,
  findSlideLayout,
  getEffectiveColorMap,
  getSlides,
  getSlideXmlString,
  getTableCell,
  inches,
  loadPresentation,
  resolveDrawingColor,
  savePresentation,
  setTableCellTextFormat,
} from '../src/api/index.ts';
import { parseXml } from '../src/internal/xml/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';

const blankFixture = (): string =>
  fileURLToPath(new URL('./fixtures/minimal/blank.pptx', import.meta.url));

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);
const encode = (s: string): Uint8Array => new TextEncoder().encode(s);

// ---------------------------------------------------------------------------
// a. resolveDrawingColor + clrMap
// ---------------------------------------------------------------------------

describe('fn API: resolveDrawingColor with clrMap param', () => {
  const theme = {
    name: 'Test',
    dark1: '#000000',
    light1: '#FFFFFF',
    dark2: '#222222',
    light2: '#EEEEEE',
    accent1: '#4472C4',
    accent2: '#ED7D31',
    accent3: '#A9D18E',
    accent4: '#8064A2',
    accent5: '#4BACC6',
    accent6: '#F79646',
    hyperlink: '#0000FF',
    followedHyperlink: '#800080',
  };

  it('WITHOUT clrMap, schemeClr tx1 resolves to dark1', () => {
    const el = parseXml(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="tx1"/>`,
    ).root;
    // tx1 directly maps to dark1 in SCHEME_TOKEN_TO_THEME_KEY
    expect(resolveDrawingColor(el, theme)).toBe('#000000');
  });

  it('WITH clrMap {tx1:"lt1"}, schemeClr tx1 resolves to light1', () => {
    const el = parseXml(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="tx1"/>`,
    ).root;
    // clrMap remaps tx1 → lt1 before the theme lookup
    expect(resolveDrawingColor(el, theme, { tx1: 'lt1' })).toBe('#FFFFFF');
  });

  it('WITH clrMap {tx1:"dk1"}, schemeClr tx1 still resolves to dark1', () => {
    const el = parseXml(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="tx1"/>`,
    ).root;
    // Standard map: tx1 → dk1 → dark1
    expect(resolveDrawingColor(el, theme, { tx1: 'dk1' })).toBe('#000000');
  });

  it('clrMap null behaves like no clrMap', () => {
    const el = parseXml(
      `<a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="tx1"/>`,
    ).root;
    expect(resolveDrawingColor(el, theme, null)).toBe('#000000');
  });

  it('non-scheme token (srgbClr) is unaffected by clrMap', () => {
    const el = parseXml(
      `<a:srgbClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="FF0000"/>`,
    ).root;
    expect(resolveDrawingColor(el, theme, { tx1: 'lt1' })).toBe('#FF0000');
  });
});

// ---------------------------------------------------------------------------
// b. getEffectiveColorMap on a createPresentation deck
// ---------------------------------------------------------------------------

describe('fn API: getEffectiveColorMap', () => {
  it('returns the standard map for a createPresentation deck', async () => {
    const pres = createPresentation();
    // createPresentation() creates the blank scaffold with layouts but no slides.
    // We must add a slide to get a SlideData with a valid master chain.
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const clrMap = getEffectiveColorMap(slide);
    // Standard map: bg1→lt1, tx1→dk1
    expect(clrMap['tx1']).toBe('dk1');
    expect(clrMap['bg1']).toBe('lt1');
    expect(clrMap['tx2']).toBe('dk2');
    expect(clrMap['bg2']).toBe('lt2');
    // Accents are identity in the standard map
    expect(clrMap['accent1']).toBe('accent1');
    expect(clrMap['accent2']).toBe('accent2');
  });

  it('returns standard map for a slide added to blank.pptx', async () => {
    const pres = await loadPresentation(await readFile(blankFixture()));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const clrMap = getEffectiveColorMap(slide);
    expect(clrMap['tx1']).toBe('dk1');
    expect(clrMap['bg1']).toBe('lt1');
  });
});

// ---------------------------------------------------------------------------
// c. addSlideTable bakes body-text color onto cell runs
// ---------------------------------------------------------------------------

describe('fn API: addSlideTable — baked color', () => {
  it('bakes 000000 onto each cell run in a createPresentation deck', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [
        ['Header A', 'Header B'],
        ['Value 1', 'Value 2'],
      ],
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    // Every run should have the baked solidFill with srgbClr val="000000"
    expect(xml).toContain('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>');
    // The rPr elements must contain the fill (not just the bare <a:rPr lang="en-US"/>)
    expect(xml).toContain('lang="en-US"');
    // There should be no cell run without the color baked in
    // (rPr immediately followed by a closing tag would be the bare form)
    expect(xml).not.toMatch(/<a:rPr lang="en-US"\/>/);
  });

  it('setTableCellTextFormat can override the baked color', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const tbl = addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(2),
      rows: [['Cell A', 'Cell B']],
    });
    // Override the first cell's color to red
    const cell = getTableCell(tbl, 0, 0);
    setTableCellTextFormat(cell, { color: '#FF0000' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const xml = getSlideXmlString(getSlides(reloaded).at(-1)!);
    // The override should be present
    expect(xml).toContain('FF0000');
  });
});

// ---------------------------------------------------------------------------
// d. addSlideChart bakes body-text color onto axis labels / legend / dLbls
// ---------------------------------------------------------------------------

describe('fn API: addSlideChart — baked color', () => {
  it('bakes 000000 on catAx and valAx txPr in a createPresentation deck', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [{ name: 'Revenue', values: [10, 20, 30] }],
      },
    });
    const bytes = await savePresentation(pres);
    const pkg = _internalPackageOf(await loadPresentation(bytes));
    const chartPart = pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml');
    expect(chartPart).toBeDefined();
    const chartXml = decode(chartPart!.data);
    // Both catAx and valAx must have a txPr with the baked color
    expect(chartXml).toContain('<c:catAx>');
    expect(chartXml).toContain('<c:valAx>');
    // The baked color should appear in the chart XML
    expect(chartXml).toContain('val="000000"');
    // Specifically in a defRPr inside a txPr
    expect(chartXml).toContain('<a:defRPr>');
    expect(chartXml).toContain('<a:solidFill>');
    expect(chartXml).toContain('<a:srgbClr val="000000"/>');
  });

  it('bakes 000000 on legend and dLbls when present', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        legend: { position: 'b' },
        dataLabels: {
          showValue: true,
          showCategory: false,
          showSeriesName: false,
          showPercent: false,
        },
      },
    });
    const bytes = await savePresentation(pres);
    const pkg = _internalPackageOf(await loadPresentation(bytes));
    const chartXml = decode(pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml')!.data);
    // Legend txPr and dLbls txPr should both contain the baked color
    expect(chartXml).toContain('<c:legend>');
    expect(chartXml).toContain('<c:dLbls>');
    // The baked color should appear multiple times (once per axis, once for legend, once for dLbls)
    const occurrences = (chartXml.match(/srgbClr val="000000"/g) ?? []).length;
    expect(occurrences).toBeGreaterThanOrEqual(3);
  });

  it('authored axis color is preserved over the baked default', async () => {
    const pres = createPresentation();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['A', 'B'],
        series: [{ name: 'X', values: [1, 2] }],
        categoryAxisLabelStyle: { color: '#FF0000' },
      },
    });
    const bytes = await savePresentation(pres);
    const pkg = _internalPackageOf(await loadPresentation(bytes));
    const chartXml = decode(pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml')!.data);
    // The authored color should be present
    expect(chartXml).toContain('FF0000');
    // The baked default (000000) should still appear for valAx
    expect(chartXml).toContain('000000');
  });
});

// ---------------------------------------------------------------------------
// e. Inverted clrMap end-to-end
// ---------------------------------------------------------------------------

describe('fn API: inverted clrMap end-to-end', () => {
  // Helper: build a createPresentation() deck and rewrite the slide master's
  // <p:clrMap> to the inverted form (bg1="dk1" tx1="lt1" ...).
  // Returns the modified presentation with a Blank slide already added.
  const buildInvertedDeck = async (): Promise<ReturnType<typeof createPresentation>> => {
    const pres = createPresentation();
    const pkg = _internalPackageOf(pres);

    // Find the slide master part
    const masterPart = pkg.parts.find((p) => p.name.startsWith('/ppt/slideMasters/slideMaster'));
    if (!masterPart) throw new Error('slide master not found');

    // Rewrite the <p:clrMap> attributes to the inverted form
    const xml = decode(masterPart.data);
    const invertedXml = xml.replace(
      /(<p:clrMap\s+)bg1="lt1"\s+tx1="dk1"\s+bg2="lt2"\s+tx2="dk2"/,
      '$1bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"',
    );
    masterPart.data = encode(invertedXml);

    // Add a slide so the deck has a SlideData we can call getEffectiveColorMap on
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    addSlide(pres, { layout });

    return pres;
  };

  it('getEffectiveColorMap reports tx1:"lt1" after inverting the master clrMap', async () => {
    const pres = await buildInvertedDeck();
    // Save and reload so the package is fresh, then check the effective map
    const reloaded = await loadPresentation(await savePresentation(pres));
    // The reloaded deck has the slide we added in buildInvertedDeck
    const slides = getSlides(reloaded);
    expect(slides.length).toBeGreaterThan(0);
    const slide = slides[0]!;
    const clrMap = getEffectiveColorMap(slide);
    expect(clrMap['tx1']).toBe('lt1');
    expect(clrMap['bg1']).toBe('dk1');
  });

  it('addSlideTable bakes FFFFFF (light1) when tx1 maps to lt1', async () => {
    const pres = await buildInvertedDeck();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideTable(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(6),
      h: inches(2),
      rows: [
        ['Cell 1', 'Cell 2'],
        ['Cell 3', 'Cell 4'],
      ],
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    // With inverted map: tx1 → lt1 → light1 = #FFFFFF
    expect(xml).toContain('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>');
    // Must NOT bake the standard 000000 (that would be the wrong color)
    expect(xml).not.toContain('val="000000"');
  });

  it('addSlideChart bakes FFFFFF on catAx when tx1 maps to lt1', async () => {
    const pres = await buildInvertedDeck();
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    addSlideChart(slide, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Jan', 'Feb'],
        series: [{ name: 'Sales', values: [100, 200] }],
      },
    });
    const bytes = await savePresentation(pres);
    const pkg = _internalPackageOf(await loadPresentation(bytes));
    const chartXml = decode(pkg.parts.find((p) => p.name === '/ppt/charts/chart1.xml')!.data);
    // The baked color should be FFFFFF, not 000000
    expect(chartXml).toContain('FFFFFF');
    expect(chartXml).not.toContain('val="000000"');
  });

  it('renderSlideToSvg flips the scheme-colored master background per the clrMap', async () => {
    // The blank-deck master paints its background with schemeClr bg1. The first
    // full-canvas <rect> in the SVG is that background. Standard map: bg1 → lt1 →
    // light1 = white. Inverted map: bg1 → dk1 → dark1 = black. The preview must
    // resolve through the effective map, so the fill flips between the two decks.
    const firstRectFill = (svg: string): string | undefined =>
      svg.match(/<rect width="[0-9.]+" height="[0-9.]+" fill="(#[0-9A-Fa-f]{6})"\/>/)?.[1];

    const std = createPresentation();
    const stdLayout = findSlideLayout(std, 'Blank');
    if (!stdLayout) throw new Error('expected Blank layout');
    const stdSlide = addSlide(std, { layout: stdLayout });
    expect(firstRectFill(renderSlideToSvg(std, stdSlide))).toBe('#FFFFFF');

    const inv = await buildInvertedDeck();
    const invSlide = getSlides(inv)[0]!;
    expect(firstRectFill(renderSlideToSvg(inv, invSlide))).toBe('#000000');
  });
});
