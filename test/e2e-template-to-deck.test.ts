// End-to-end proof: canonical template-to-deck authoring path.
//
// Scenario: a team receives layout-decoration.pptx (a company template with a
// logo and a "TEMPLATE" bar injected into the layout, and a footer bar on the
// master) and needs to produce a new presentation from it — keeping the
// masters, layouts, and theme intact but replacing the template's own slides
// with fresh content.
//
// This file IS documentation for that path. A reader who wants to know how to
// use pptx-kit for template-to-deck authoring should be able to read top to
// bottom and understand every step.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideChart,
  addSlideImage,
  addSlideTable,
  addSlideTextBox,
  duplicateSlide,
  findSlidePlaceholder,
  findSlideLayout,
  getShapeText,
  getSlideLayout,
  getSlideLayoutName,
  getSlideLayouts,
  getSlideSize,
  getSlides,
  getSlideText,
  importSlide,
  inches,
  loadPresentation,
  removeSlide,
  replaceTokensInPresentation,
  savePresentation,
  setShapeText,
} from '../src/api/index.ts';
import { renderSlideToSvg } from '../packages/preview/src/index.ts';
import { renderSlideToRgba } from '../packages/preview/src/node.ts';
import { buildPng } from './lib/build-png.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';
import { expectSemanticallyEqual } from './lib/semantic-equal.ts';
import { textContentOf } from './lib/svg-query.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// ── Helpers ──────────────────────────────────────────────────────────────────

const decode = (b: Uint8Array): string => new TextDecoder().decode(b);

// xmllint is optional — schema tests skip cleanly if it's not installed.
const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

// Known layout names that must survive template adoption. They come from the
// python-pptx default theme, which layout-decoration.pptx inherits.
const TEMPLATE_LAYOUT_NAMES = ['Title Slide', 'Title and Content', 'Title Only', 'Blank'] as const;

// The deck we build has exactly this many slides.
const EXPECTED_SLIDE_COUNT = 6;

// The 32×32 solid-blue PNG we embed on the media slide.
const BLUE_PNG = buildPng(32, 32, [0x3b, 0x82, 0xf6]);

// Token values for slide 2's token-replacement pass.
const TOKEN_VALUES = { company: 'Acme Corp', year: '2026' } as const;

// ── Main scenario ────────────────────────────────────────────────────────────

describe('E2E: template-to-deck', () => {
  // Shared state: built once, reused across all `it` cases below.
  // Vitest runs each `it` in serial within a `describe`, so this is safe.
  let deckBytes: Uint8Array;
  let reloaded: Awaited<ReturnType<typeof loadPresentation>>;

  it('builds a 6-slide deck from the company template', async () => {
    // ── Step 1: load the template ─────────────────────────────────────────────
    const template = await loadPresentation(await readFile(fixture('layout-decoration.pptx')));

    // The template fixture starts with one placeholder slide that was used to
    // author the master/layout decoration. We drop it — users of a real
    // template expect zero slides, not the author's example.
    const templateSlides = getSlides(template);
    expect(templateSlides.length).toBe(1);
    for (const s of templateSlides) removeSlide(template, s);
    expect(getSlides(template).length).toBe(0);

    // ── Step 2: add content slides from the template's layouts ────────────────

    // Slide 1: Cover slide. Title Slide layout → ctrTitle + subTitle.
    const titleLayout = findSlideLayout(template, 'Title Slide');
    if (!titleLayout) throw new Error('Title Slide layout not found in template');
    const slide1 = addSlide(template, { layout: titleLayout });
    const ctrTitle = findSlidePlaceholder(slide1, 'ctrTitle');
    const subTitle = findSlidePlaceholder(slide1, 'subTitle');
    if (!ctrTitle) throw new Error('ctrTitle placeholder missing on Title Slide');
    if (!subTitle) throw new Error('subTitle placeholder missing on Title Slide');
    setShapeText(ctrTitle, 'pptx-kit E2E');
    setShapeText(subTitle, 'template-to-deck proof');

    // Slide 2: Token-replacement slide. The body carries {{company}} / {{year}}
    // tokens that replaceTokensInPresentation will fill in. Deliberately
    // authored on a separate slide from slide 1 to prove multi-slide
    // token replacement works in a single call.
    const contentLayout = findSlideLayout(template, 'Title and Content');
    if (!contentLayout) throw new Error('Title and Content layout not found');
    const slide2 = addSlide(template, { layout: contentLayout });
    const s2Title = findSlidePlaceholder(slide2, 'title');
    const s2Body = findSlidePlaceholder(slide2, 'body');
    if (!s2Title) throw new Error('title placeholder missing on Title and Content');
    setShapeText(s2Title, '{{company}} Report');
    if (s2Body) setShapeText(s2Body, '{{year}} edition');

    const replacedCount = replaceTokensInPresentation(template, TOKEN_VALUES);
    // Two placeholders carry tokens — title and body of slide 2.
    expect(replacedCount).toBeGreaterThanOrEqual(2);

    // Slide 3: Media slide. Title Only layout + a free-form image + a caption
    // text box added as independent shapes (not as placeholders). This proves
    // addSlideImage and addSlideTextBox work alongside placeholder-based slides
    // in the same deck.
    const titleOnlyLayout = findSlideLayout(template, 'Title Only');
    if (!titleOnlyLayout) throw new Error('Title Only layout not found');
    const slide3 = addSlide(template, { layout: titleOnlyLayout });
    const s3Title = findSlidePlaceholder(slide3, 'title');
    if (s3Title) setShapeText(s3Title, 'Media slide');
    addSlideImage(slide3, BLUE_PNG, {
      x: inches(1),
      y: inches(1.5),
      w: inches(2),
      h: inches(2),
    });
    addSlideTextBox(slide3, {
      x: inches(1),
      y: inches(3.7),
      w: inches(2),
      h: inches(0.4),
      text: 'Blue square (32×32 px)',
    });

    // Slide 4: Data slide. Blank layout + a column chart + a summary table.
    // Both are independent shapes, not placeholders.
    const blankLayout = findSlideLayout(template, 'Blank');
    if (!blankLayout) throw new Error('Blank layout not found');
    const slide4 = addSlide(template, { layout: blankLayout });
    addSlideChart(slide4, {
      x: inches(0.5),
      y: inches(0.5),
      w: inches(5),
      h: inches(3.5),
      spec: {
        kind: 'column',
        title: 'Q1–Q4 Revenue',
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [{ name: 'Revenue', values: [100, 120, 90, 140] }],
      },
    });
    addSlideTable(slide4, {
      x: inches(5.7),
      y: inches(0.5),
      w: inches(3.8),
      h: inches(2),
      rows: [
        ['Quarter', 'Revenue'],
        ['Q1', '$100k'],
        ['Q2', '$120k'],
        ['Q3', '$90k'],
        ['Q4', '$140k'],
      ],
    });

    // Slide 5: Duplicate of the cover slide. duplicateSlide copies the slide
    // XML independently — mutating the duplicate must not affect the original.
    const slide5 = duplicateSlide(template, slide1);
    const dupTitle = findSlidePlaceholder(slide5, 'ctrTitle');
    if (dupTitle) setShapeText(dupTitle, 'pptx-kit E2E (copy)');

    // Slide 6: Cross-deck import from one-text-slide.pptx.
    // importSlide brings the source slide's shapes into the target deck,
    // re-binding it to the supplied layout. This exercises cross-deck
    // composition — e.g. combining a slide from one source deck with a
    // different template.
    const source = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const sourceSlide = getSlides(source)[0]!;
    importSlide(template, sourceSlide, titleOnlyLayout);

    // ── Step 3: assert pre-save state ─────────────────────────────────────────
    expect(getSlides(template).length).toBe(EXPECTED_SLIDE_COUNT);

    // ── Step 4: save → reload ─────────────────────────────────────────────────
    deckBytes = await savePresentation(template);
    reloaded = await loadPresentation(deckBytes);
  });

  it('reloaded deck has the right slide count and text', () => {
    // Guard: depends on the build step above.
    if (!reloaded) return;

    const slides = getSlides(reloaded);
    expect(slides.length).toBe(EXPECTED_SLIDE_COUNT);

    // Slide 1: cover title survived round-trip.
    const s1CtrTitle = findSlidePlaceholder(slides[0]!, 'ctrTitle');
    expect(s1CtrTitle && getShapeText(s1CtrTitle)).toBe('pptx-kit E2E');

    // Slide 2: tokens were replaced before saving.
    const s2Title = findSlidePlaceholder(slides[1]!, 'title');
    expect(s2Title && getShapeText(s2Title)).toBe('Acme Corp Report');
    const s2Body = findSlidePlaceholder(slides[1]!, 'body');
    expect(s2Body && getShapeText(s2Body)).toBe('2026 edition');

    // Slide 5: duplicate got its own independent title mutation.
    const s5Title = findSlidePlaceholder(slides[4]!, 'ctrTitle');
    expect(s5Title && getShapeText(s5Title)).toBe('pptx-kit E2E (copy)');
    // Original slide 1 text unchanged.
    const s1Still = findSlidePlaceholder(slides[0]!, 'ctrTitle');
    expect(s1Still && getShapeText(s1Still)).toBe('pptx-kit E2E');

    // Slide 6: imported slide carries "Hello, OOXML" from the source deck.
    // getSlideText collapses all shape text into a single string.
    expect(getSlideText(slides[5]!)).toContain('Hello, OOXML');
  });

  it('reloaded deck preserves the template master/layouts', () => {
    if (!reloaded) return;

    const layoutNames = getSlideLayouts(reloaded).map(getSlideLayoutName);

    // Every layout present in the original template must survive, regardless of
    // whether a slide uses it. Layouts are part of the presentation's theme
    // identity and must not be stripped on save.
    for (const name of TEMPLATE_LAYOUT_NAMES) {
      expect(layoutNames).toContain(name);
    }

    // Every slide is bound to a layout from the reloaded deck (not a dangling ref).
    for (const slide of getSlides(reloaded)) {
      const layout = getSlideLayout(slide);
      expect(layout).not.toBeNull();
      expect(getSlideLayoutName(layout!)).toBeTruthy();
    }
  });

  skipIfNoXmllint('all emitted slide XML parts are schema-valid', () => {
    if (!reloaded) return;

    const pkg = _internalPackageOf(reloaded);
    for (const part of pkg.parts) {
      if (part.contentType.endsWith('slide+xml')) {
        expectSchemaValid(decode(part.data), 'pml');
      }
    }
  });

  skipIfNoXmllint('presentation.xml is schema-valid after the build', () => {
    if (!reloaded) return;

    const pkg = _internalPackageOf(reloaded);
    const presPart = pkg.parts.find((p) => p.contentType.endsWith('presentation.main+xml'));
    expect(presPart).toBeDefined();
    expectSchemaValid(decode(presPart!.data), 'pml');
  });

  it('re-save is semantically identical to the first save (round-trip stable)', async () => {
    if (!reloaded) return;

    const bytes2 = await savePresentation(reloaded);
    const reloaded2 = await loadPresentation(bytes2);

    // The semantic comparator tolerates ZIP compression non-determinism, XML
    // whitespace differences, and attribute-order variation. It fails on any
    // structural difference: different parts, different content types, or
    // different XML element trees.
    expectSemanticallyEqual(_internalPackageOf(reloaded), _internalPackageOf(reloaded2));
  });

  it('renderSlideToSvg (foreignObject mode): no rendering-failure fallbacks, text present', () => {
    if (!reloaded) return;

    const slides = getSlides(reloaded);
    for (const slide of slides) {
      const svg = renderSlideToSvg(reloaded, slide, { textLayout: 'foreignObject' });

      // A data-pptx-fallback marker means the renderer could not produce real
      // content for a shape we authored. Nothing in this deck uses an
      // unsupported feature, so no marker of any kind may appear.
      expect(svg).not.toContain('data-pptx-fallback');
    }

    // Slide 1 cover title text is present in the rendered output.
    const s1Svg = renderSlideToSvg(reloaded, slides[0]!, { textLayout: 'foreignObject' });
    expect(textContentOf(s1Svg)).toContain('pptx-kit E2E');

    // Slide 6 imported text is present.
    const s6Svg = renderSlideToSvg(reloaded, slides[5]!, { textLayout: 'foreignObject' });
    expect(textContentOf(s6Svg)).toContain('Hello, OOXML');
  });

  it('renderSlideToSvg (svg text mode): no rendering-failure fallbacks, text present', () => {
    if (!reloaded) return;

    const slides = getSlides(reloaded);
    for (const slide of slides) {
      const svg = renderSlideToSvg(reloaded, slide, { textLayout: 'svg' });
      expect(svg).not.toContain('data-pptx-fallback');
    }

    // In svg text mode the renderer emits <text> elements; the cover title
    // must appear as actual text content, not as a foreignObject fallback.
    const s1Svg = renderSlideToSvg(reloaded, slides[0]!, { textLayout: 'svg' });
    expect(textContentOf(s1Svg)).toContain('pptx-kit E2E');
  });

  it('renderSlideToRgba (Node): first two slides produce non-blank rasters at 480 px', () => {
    if (!reloaded) return;

    const slides = getSlides(reloaded);
    const slideSize = getSlideSize(reloaded);
    if (!slideSize) throw new Error('getSlideSize returned null');

    // Width 480 keeps the test fast; height follows the slide's aspect ratio.
    const WIDTH = 480;

    for (const slide of [slides[0]!, slides[1]!]) {
      const { image } = renderSlideToRgba(reloaded, slide, { width: WIDTH });
      expect(image.width).toBe(WIDTH);
      // Height rounds to match the slide aspect ratio.
      const expectedHeight = Math.round((WIDTH * slideSize.height) / slideSize.width);
      expect(Math.abs(image.height - expectedHeight)).toBeLessThanOrEqual(1);
      expect(image.data.length).toBe(image.width * image.height * 4);

      // A non-blank frame has at least one pixel that isn't pure white.
      // Slides 1 and 2 carry text so the rendered pixels include ink.
      let nonWhitePixels = 0;
      for (let i = 0; i < image.data.length; i += 4) {
        if (image.data[i] !== 255 || image.data[i + 1] !== 255 || image.data[i + 2] !== 255) {
          nonWhitePixels++;
        }
      }
      expect(nonWhitePixels).toBeGreaterThan(0);
    }
  });
});
