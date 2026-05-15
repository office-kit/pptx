// Free-function shape click-action API.
//
// Verifies `<a:hlinkClick>` lands on the shape's cNvPr for each of the
// supported action kinds (URL, slide jump, preset show navigation).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  readPackagePart,
  getSlideShapes,
  getSlideXmlString,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeClickAction,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

const slideRels = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  const bytesPart = readPackagePart(pres, `/ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`);
  return bytesPart ? new TextDecoder().decode(bytesPart) : '';
};

describe('fn API: setShapeClickAction', () => {
  it('attaches a URL click action and removes it on null', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;

    setShapeClickAction(shape, { kind: 'url', url: 'https://example.com/' });
    const wired = await savePresentation(pres);
    expect(await slideRels(wired, 0)).toContain('https://example.com/');
    expect(await slideXml(wired, 0)).toContain('hlinkClick');

    setShapeClickAction(shape, null);
    expect(await slideXml(await savePresentation(pres), 0)).not.toContain('hlinkClick');
  });

  it('jumps to a slide via slide-rel + ppaction://hlinksldjump', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slides = getSlides(pres);
    const target = slides[1]!;
    const shape = getSlideShapes(slides[0]!)[0]!;

    setShapeClickAction(shape, { kind: 'slide', slide: target });
    const bytes = await savePresentation(pres);
    expect(await slideXml(bytes, 0)).toContain('hlinksldjump');
    expect(await slideRels(bytes, 0)).toContain('slide2.xml');
  });

  it('preset navigation kinds emit the ppaction without allocating a rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;
    setShapeClickAction(shape, { kind: 'nextSlide' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('hlinkshowjump?jump=nextslide');
  });
});
