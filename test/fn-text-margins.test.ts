// setShapeTextMargins — internal padding on bodyPr.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideTextBox,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeTextMargins,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeTextMargins', () => {
  it('writes lIns/tIns/rIns/bIns attributes on bodyPr', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2), text: 'with padding',
    });
    setShapeTextMargins(tb, { left: 0, top: 50000, right: 100000, bottom: 25000 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toMatch(/<a:bodyPr[^>]*lIns="0"/);
    expect(xml).toMatch(/<a:bodyPr[^>]*tIns="50000"/);
    expect(xml).toMatch(/<a:bodyPr[^>]*rIns="100000"/);
    expect(xml).toMatch(/<a:bodyPr[^>]*bIns="25000"/);
  });

  it('only writes specified sides and replaces them on subsequent calls', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const tb = addSlideTextBox(slide, {
      x: inches(0), y: inches(0), w: inches(3), h: inches(2), text: 'A',
    });
    setShapeTextMargins(tb, { left: 1000 });
    setShapeTextMargins(tb, { left: 2000 });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('lIns="2000"');
    expect(xml).not.toContain('lIns="1000"');
  });
});
