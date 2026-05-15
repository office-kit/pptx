// Shape outline (a:ln) authoring.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  pt,
  setShapeFill,
  setShapeNoStroke,
  setShapeStroke,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

describe('L3: setShapeStroke', () => {
  it('sets line color and width', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1), y: inches(1), w: inches(4), h: inches(1), text: 'Bordered',
    });
    setShapeStroke(box, { color: '#FF8800', widthEmu: pt(2) });

    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('val="FF8800"');
    expect(xml).toMatch(/<a:ln[^>]*w="25400"/); // 2pt = 25400 EMU
  });

  it('toggles to noStroke', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1), y: inches(1), w: inches(4), h: inches(1), text: 't',
    });
    setShapeStroke(box, { color: '#000000', widthEmu: pt(1) });
    setShapeNoStroke(box);
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<a:ln');
    expect(xml).toContain('<a:noFill/>');
  });

  skipIfNoXmllint('outlined shape validates against pml.xsd', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const box = addSlideTextBox(slide, {
      x: inches(1), y: inches(1), w: inches(4), h: inches(1), text: 't',
    });
    setShapeStroke(box, { color: 'accent3', widthEmu: pt(1.5) });
    setShapeFill(box, '#FFFFFF');
    setShapeTextFormat(box, { color: '#222222', size: 14 });

    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
