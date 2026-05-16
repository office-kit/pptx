// Paragraph alignment + hyperlinks.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideTextBox,
  findSlideLayout,
  getShapeHyperlink,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  setShapeAlignment,
  setShapeHyperlink,
  setShapeText,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const newBox = async (text = 'x') => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected Blank layout');
  const slide = addSlide(pres, { layout });
  const box = addSlideTextBox(slide, {
    x: inches(1), y: inches(1), w: inches(8), h: inches(2), text,
  });
  return { pres, box };
};

describe('L3: paragraph alignment', () => {
  it('writes algn="ctr" on every paragraph for setShapeAlignment("center")', async () => {
    const { pres, box } = await newBox('left line\ncenter line');
    setShapeText(box, 'first\nsecond');
    setShapeAlignment(box, 'center');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect((xml.match(/algn="ctr"/g) ?? []).length).toBe(2);
  });

  it('accepts ECMA-376 raw tokens directly', async () => {
    const { pres, box } = await newBox();
    setShapeAlignment(box, 'just');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('algn="just"');
  });

  skipIfNoXmllint('aligned shape validates against pml.xsd', async () => {
    const { pres, box } = await newBox('a\nb');
    setShapeAlignment(box, 'center');
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});

describe('L3: hyperlinks', () => {
  it('adds a hyperlink rel and a:hlinkClick rPr child', async () => {
    const { pres, box } = await newBox('Click me');
    setShapeHyperlink(box, 'https://example.com/docs');
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<a:hlinkClick');
    expect(xml).toMatch(/r:id="rId\d+"/);
    expect(getShapeHyperlink(box)).toBe('https://example.com/docs');
  });

  it('reuses an existing hyperlink rel for the same URL', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(pres, 'Blank');
    if (!layout) throw new Error('expected Blank layout');
    const slide = addSlide(pres, { layout });
    const a = addSlideTextBox(slide, {
      x: inches(1), y: inches(1), w: inches(4), h: inches(1), text: 'A',
    });
    const b = addSlideTextBox(slide, {
      x: inches(1), y: inches(3), w: inches(4), h: inches(1), text: 'B',
    });
    setShapeHyperlink(a, 'https://example.com');
    setShapeHyperlink(b, 'https://example.com');

    // Both shapes report the same URL; the rels-allocation reuse is
    // implicit (no duplicate target).
    const urls: string[] = [];
    for (const sh of getSlideShapes(getSlides(pres).at(-1)!)) {
      const u = getShapeHyperlink(sh);
      if (u !== null) urls.push(u);
    }
    expect(urls).toEqual(['https://example.com', 'https://example.com']);
  });

  it('removes the hlinkClick when setShapeHyperlink(null) is called', async () => {
    const { pres, box } = await newBox();
    setShapeHyperlink(box, 'https://example.com');
    setShapeHyperlink(box, null);
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).not.toContain('<a:hlinkClick');
  });

  skipIfNoXmllint('hyperlink output validates against pml.xsd', async () => {
    const { pres, box } = await newBox('docs');
    setShapeHyperlink(box, 'https://example.com');
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
