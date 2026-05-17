// Text run formatting: font, size, color, bold, italic, underline.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getSlideText,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeText,
  setShapeTextFormat,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const newBox = async (text = 'Styled text') => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected Blank layout');
  const slide = addSlide(pres, { layout });
  const box = addSlideTextBox(slide, {
    x: inches(1),
    y: inches(1),
    w: inches(4),
    h: inches(1),
    text,
  });
  return { pres, box };
};

describe('L3: setShapeTextFormat', () => {
  it('sets size, bold, italic, underline, font, and color on all runs', async () => {
    const { pres, box } = await newBox('Styled text');
    setShapeTextFormat(box, {
      font: 'Calibri',
      size: 24,
      color: '#3366CC',
      bold: true,
      italic: true,
      underline: true,
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('sz="2400"');
    expect(xml).toContain('b="1"');
    expect(xml).toContain('i="1"');
    expect(xml).toContain('u="sng"');
    expect(xml).toContain('<a:latin typeface="Calibri"/>');
    expect(xml).toContain('val="3366CC"');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideText(getSlides(reloaded)[0]!)).toContain('Styled text');
  });

  it('partial updates compose without losing prior formatting', async () => {
    const { pres, box } = await newBox('Compose');
    setShapeTextFormat(box, { bold: true });
    setShapeTextFormat(box, { color: '#FF0000' });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('b="1"');
    expect(xml).toContain('val="FF0000"');
  });

  it('accepts scheme-color tokens', async () => {
    const { pres, box } = await newBox('Theme color');
    setShapeTextFormat(box, { color: 'accent1' });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<a:schemeClr val="accent1"/>');
  });

  it('rejects unrecognized color values', async () => {
    const { box } = await newBox('t');
    expect(() => setShapeTextFormat(box, { color: 'pumpkin spice' })).toThrow(/color/);
  });

  skipIfNoXmllint('formatted output validates against pml.xsd', async () => {
    const { pres, box } = await newBox('Validates with format');
    setShapeTextFormat(box, {
      font: 'Helvetica',
      size: 18.5,
      color: '#222222',
      bold: true,
      italic: false,
      underline: 'dbl',
    });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });

  it('applies formatting before setShapeText and the format survives', async () => {
    const { pres, box } = await newBox('Initial');
    setShapeTextFormat(box, { bold: true, color: '#00AA00' });
    setShapeText(box, 'Replaced text');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('b="1"');
    expect(xml).toContain('val="00AA00"');
    expect(xml).toContain('Replaced text');
  });
});
