// setShapeStrokeArrow — arrowheads on `<a:ln>`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideLine,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeStrokeArrow,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

describe('fn API: setShapeStrokeArrow', () => {
  it('writes a tailEnd of the configured type / size', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(3), y: inches(2) },
      color: '#000000',
      widthEmu: 12700,
    });
    setShapeStrokeArrow(line, 'tail', { type: 'triangle', width: 'med', length: 'lg' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:tailEnd');
    expect(xml).toContain('type="triangle"');
    expect(xml).toContain('w="med"');
    expect(xml).toContain('len="lg"');
  });

  it('reapplying replaces the prior arrowhead on the same end', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(3), y: inches(2) },
    });
    setShapeStrokeArrow(line, 'tail', { type: 'triangle' });
    setShapeStrokeArrow(line, 'tail', { type: 'arrow' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('type="arrow"');
    expect(xml).not.toContain('type="triangle"');
  });

  it('head and tail are independent', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const line = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(3), y: inches(2) },
    });
    setShapeStrokeArrow(line, 'head', { type: 'diamond' });
    setShapeStrokeArrow(line, 'tail', { type: 'oval' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:headEnd');
    expect(xml).toContain('type="diamond"');
    expect(xml).toContain('<a:tailEnd');
    expect(xml).toContain('type="oval"');
  });
});
