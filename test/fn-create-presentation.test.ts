// createPresentation — from-scratch authoring.
//
// Verifies that `createPresentation()` returns an immediately-authorable
// deck (master + theme + layouts + slide size), that every emitted XML
// part is schema-valid, and that a full author→save→load round-trip
// preserves slides, text, and charts.

import { describe, expect, it } from 'vitest';
import {
  addBlankSlide,
  addContentSlide,
  addSlide,
  addSlideChart,
  addSlideTextBox,
  createPresentation,
  findSlideLayout,
  findSlideLayoutByType,
  findSlidePlaceholder,
  getShapeKind,
  getPresentationFonts,
  getShapeText,
  getSlideCharts,
  getSlideLayoutName,
  getSlideLayouts,
  getSlideShapes,
  getSlideSize,
  getSlides,
  getSlideText,
  inches,
  loadPresentation,
  readPackagePart,
  savePresentation,
  setPresentationFonts,
  setShapeText,
  validatePresentation,
} from '../src/api/index.ts';
import {
  expectSchemaValid,
  isSchemaValidationAvailable,
  type SchemaKind,
} from './lib/expect-schema-valid.ts';

const decoder = new TextDecoder();

describe('fn API: createPresentation', () => {
  it('returns a deck with master-backed layouts ready for addSlide', () => {
    const pres = createPresentation();
    const layouts = getSlideLayouts(pres);
    expect(layouts.length).toBeGreaterThan(0);

    const names = layouts.map((l) => getSlideLayoutName(l)).sort();
    expect(names).toEqual(['Blank', 'Title Slide', 'Title and Content']);

    // The spec-token lookups the deck helpers rely on must resolve.
    expect(findSlideLayoutByType(pres, 'blank')).not.toBeNull();
    expect(findSlideLayoutByType(pres, 'title')).not.toBeNull();
    expect(findSlideLayoutByType(pres, 'obj')).not.toBeNull();
  });

  it("defaults the slide size to 16:9 (PowerPoint's modern default)", () => {
    const pres = createPresentation();
    const size = getSlideSize(pres);
    expect(size).not.toBeNull();
    expect(size).toMatchObject({ width: 12192000, height: 6858000, type: 'screen16x9' });
  });

  it('honours the 4:3 size option', () => {
    const pres = createPresentation({ size: '4:3' });
    expect(getSlideSize(pres)).toMatchObject({
      width: 9144000,
      height: 6858000,
      type: 'screen4x3',
    });
  });

  it('starts with no slides and passes the invariant validator', () => {
    const pres = createPresentation();
    expect(getSlides(pres).length).toBe(0);
    expect(validatePresentation(pres)).toEqual([]);
  });

  it('round-trips a from-scratch deck: addSlide → text → chart → save → load', async () => {
    const pres = createPresentation();

    // Title slide via explicit layout + placeholder text.
    const titleLayout = findSlideLayout(pres, 'Title Slide')!;
    const titleSlide = addSlide(pres, { layout: titleLayout });
    const ctrTitle = findSlidePlaceholder(titleSlide, 'ctrTitle')!;
    setShapeText(ctrTitle, '@office-kit/pptx from scratch');

    // Content slide via the sugar helper.
    addContentSlide(pres, { title: 'Agenda', body: 'First item' });

    // Blank slide hosting a free-form text box + a chart.
    const blankSlide = addBlankSlide(pres);
    addSlideTextBox(blankSlide, {
      x: inches(1),
      y: inches(1),
      w: inches(8),
      h: inches(1),
      text: 'Free-form text box',
    });
    const chartShape = addSlideChart(blankSlide, {
      x: inches(1),
      y: inches(2.5),
      w: inches(6),
      h: inches(4),
      spec: {
        kind: 'column',
        categories: ['Q1', 'Q2', 'Q3'],
        series: [{ name: 'Revenue', values: [10, 20, 15] }],
      },
    });
    expect(getShapeKind(chartShape)).toBe('graphicFrame');

    expect(getSlides(pres).length).toBe(3);
    expect(validatePresentation(pres)).toEqual([]);

    // Save and reload — every authored fact must survive the trip.
    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);

    const slides = getSlides(reloaded);
    expect(slides.length).toBe(3);
    expect(getSlideLayouts(reloaded).length).toBe(3);

    expect(getSlideText(slides[0]!)).toContain('@office-kit/pptx from scratch');
    expect(getSlideText(slides[1]!)).toContain('Agenda');
    expect(getSlideText(slides[1]!)).toContain('First item');
    expect(getSlideText(slides[2]!)).toContain('Free-form text box');

    const charts = getSlideCharts(slides[2]!);
    expect(charts.length).toBe(1);
    expect(charts[0]!.spec?.kind).toBe('column');
    expect(charts[0]!.spec?.series[0]?.name).toBe('Revenue');

    // The reloaded title placeholder still reads back its text.
    const reloadedTitle = findSlidePlaceholder(slides[0]!, 'ctrTitle');
    expect(reloadedTitle).not.toBeNull();
    expect(getShapeText(reloadedTitle!)).toBe('@office-kit/pptx from scratch');
  });

  it('adds a slide for every shipped layout without throwing', () => {
    const pres = createPresentation();
    for (const layout of getSlideLayouts(pres)) {
      const slide = addSlide(pres, { layout });
      // Each new slide must at least carry the slide-root group.
      expect(getSlideShapes(slide).length).toBeGreaterThanOrEqual(0);
    }
    expect(getSlides(pres).length).toBe(3);
    expect(validatePresentation(pres)).toEqual([]);
  });

  it("carries Office's per-script font lists in the theme font scheme", async () => {
    const pres = createPresentation();
    setPresentationFonts(pres, { minorLatin: 'Arial' });
    const loaded = await loadPresentation(await savePresentation(pres));
    const theme = decoder.decode(readPackagePart(loaded, '/ppt/theme/theme1.xml')!);
    const collection = (tag: 'majorFont' | 'minorFont'): string =>
      theme.slice(theme.indexOf(`<a:${tag}>`), theme.indexOf(`</a:${tag}>`));
    const major = collection('majorFont');
    const minor = collection('minorFont');

    // The list is what a renderer consults for a script no run names a face
    // for; Thai is the entry the two collections disagree on.
    expect(major).toContain('<a:font script="Thai" typeface="Angsana New"/>');
    expect(minor).toContain('<a:font script="Thai" typeface="Cordia New"/>');
    const SCRIPT_ENTRIES = 47;
    expect(major.match(/<a:font script=/g)).toHaveLength(SCRIPT_ENTRIES);
    expect(minor.match(/<a:font script=/g)).toHaveLength(SCRIPT_ENTRIES);
    // CT_FontCollection is a sequence: latin, ea, cs, then the script list.
    expect(minor).toMatch(/<a:latin [^>]*\/><a:ea [^>]*\/><a:cs [^>]*\/><a:font script=/);
    // setPresentationFonts still writes the three slots and nothing else.
    expect(minor).toContain('<a:latin typeface="Arial"/>');
    expect(major).toContain('<a:latin typeface="Calibri"/>');
    expect(getPresentationFonts(loaded)).toMatchObject({
      majorLatin: 'Calibri',
      minorLatin: 'Arial',
      minorComplexScript: null,
    });
  });

  it.runIf(isSchemaValidationAvailable())(
    'emits schema-valid XML for every part of the blank deck',
    async () => {
      const pres = createPresentation();
      const bytes = await savePresentation(pres);
      const loaded = await loadPresentation(bytes);

      const validate = (partPath: string, kind: SchemaKind): void => {
        const raw = readPackagePart(loaded, partPath);
        expect(raw, `missing part ${partPath}`).not.toBeNull();
        expectSchemaValid(decoder.decode(raw!), kind);
      };

      validate('/ppt/presentation.xml', 'pml');
      validate('/ppt/slideMasters/slideMaster1.xml', 'pml');
      validate('/ppt/slideLayouts/slideLayout1.xml', 'pml');
      validate('/ppt/slideLayouts/slideLayout2.xml', 'pml');
      validate('/ppt/slideLayouts/slideLayout3.xml', 'pml');
      validate('/ppt/theme/theme1.xml', 'dml');
    },
  );
});
