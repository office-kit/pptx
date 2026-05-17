// Shape fill + slide background fill.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  clearSlideBackground,
  findSlideLayout,
  findSlidePlaceholder,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeFill,
  setShapeNoFill,
  setSlideBackground,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: setShapeFill', () => {
  it('sets a solid srgb fill on a text box', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 'Filled',
    });
    setShapeFill(box, '#FFCC00');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('val="FFCC00"');
    expect(xml).not.toContain('<a:noFill/>');
  });

  it('toggles to noFill and back', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 't',
    });
    setShapeFill(box, '#FF0000');
    setShapeNoFill(box);
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).not.toContain('solidFill');
    expect(xml).toContain('<a:noFill/>');
  });

  it('applies fill to a placeholder', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const title = findSlidePlaceholder(slide, 'title');
    if (title) setShapeFill(title, '#0066CC');
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('val="0066CC"');
  });

  skipIfNoXmllint('filled shape XML validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(4),
      h: inches(1),
      text: 't',
    });
    setShapeFill(box, 'accent2');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expectSchemaValid(xml, 'pml');
  });
});

describe('L3: setSlideBackground', () => {
  it('inserts <p:bg><p:bgPr><a:solidFill>... before <p:spTree>', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#003366');

    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).toContain('<p:bg>');
    expect(xml).toContain('<p:bgPr>');
    expect(xml).toContain('val="003366"');
    expect(xml.indexOf('<p:bg>')).toBeLessThan(xml.indexOf('<p:spTree>'));
  });

  it('clearSlideBackground removes any prior bg element', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#FF0000');
    clearSlideBackground(getSlides(pres)[0]!);
    const xml = getSlideXmlString(getSlides(pres)[0]!);
    expect(xml).not.toContain('<p:bg>');
  });

  skipIfNoXmllint('background-fill XML validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackground(slide, '#EFEFEF');
    expectSchemaValid(getSlideXmlString(getSlides(pres)[0]!), 'pml');
  });
});
