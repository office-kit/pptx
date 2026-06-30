// Connector lines (p:cxnSp with prstGeom prst="line").

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlide,
  addSlideLine,
  findSlideLayout,
  getShapeKind,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  pt,
} from '../src/api/index.ts';
import { expectSchemaValid, isSchemaValidationAvailable } from './lib/expect-schema-valid.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const skipIfNoXmllint = isSchemaValidationAvailable() ? it : it.skip;

const newSlide = async () => {
  const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
  const layout = findSlideLayout(pres, 'Blank');
  if (!layout) throw new Error('expected layout');
  const slide = addSlide(pres, { layout });
  return { pres, slide };
};

describe('L3: addSlideLine', () => {
  it('emits a p:cxnSp with prstGeom prst="line"', async () => {
    const { pres, slide } = await newSlide();
    const line = addSlideLine(slide, {
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(5), y: inches(3) },
      color: '#0066CC',
      widthEmu: pt(2),
    });
    expect(getShapeKind(line)).toBe('connector');

    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('prst="line"');
    expect(xml).toContain('val="0066CC"');
    expect(xml).toMatch(/<a:ln[^>]*w="25400"/);
  });

  it('writes flipH / flipV when from is to the right of / below to', async () => {
    const { pres, slide } = await newSlide();
    addSlideLine(slide, {
      from: { x: inches(5), y: inches(3) },
      to: { x: inches(1), y: inches(1) },
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('flipH="1"');
    expect(xml).toContain('flipV="1"');
  });

  it('gives an unstyled connector a visible default line style', async () => {
    const { pres, slide } = await newSlide();
    addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(1), y: inches(0) },
    });
    const xml = getSlideXmlString(getSlides(pres).at(-1)!);
    expect(xml).toContain('<p:cxnSp>');
    expect(xml).toContain('prst="line"');
    // No explicit <a:ln> (inherits width), but a <p:style><a:lnRef idx="1"> so
    // the line still renders — without it the connector would be invisible.
    expect((xml.match(/<a:ln[ >]/g) ?? []).length).toBe(0);
    expect(xml).toContain('<p:style>');
    expect(xml).toContain('<a:lnRef idx="1">');
  });

  skipIfNoXmllint('connector validates against pml.xsd', async () => {
    const { pres, slide } = await newSlide();
    addSlideLine(slide, {
      from: { x: inches(1), y: inches(1) },
      to: { x: inches(5), y: inches(3) },
      color: 'accent1',
      widthEmu: pt(1.5),
    });
    expectSchemaValid(getSlideXmlString(getSlides(pres).at(-1)!), 'pml');
  });
});
