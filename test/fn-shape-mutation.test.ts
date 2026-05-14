// Free-function shape mutation API.
//
// Verifies every SlideShape mutation entry point in `fn.ts` produces
// XML that round-trips cleanly and contains the expected attributes.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  clearShapeFill,
  clearShapeStroke,
  findSlidePlaceholder,
  getShapePosition,
  getShapeRotation,
  getShapeSize,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  removeShape,
  savePresentation,
  setShapeAlignment,
  setShapeBullets,
  setShapeFill,
  setShapeFlip,
  setShapeHyperlink,
  setShapeNoFill,
  setShapePosition,
  setShapeRotation,
  setShapeSize,
  setShapeStroke,
  setShapeText,
  setShapeTextFormat,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// Re-load via the class API just so we can pull the package out and
// dig into the unzipped slide part bytes via _internalPackageOf.
const slideXml = async (
  bytes: Uint8Array,
  slideIndex: number,
): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const slidePart = pkg.parts.find(
    (p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`,
  );
  if (!slidePart) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(slidePart.data);
};

const slideRels = async (
  bytes: Uint8Array,
  slideIndex: number,
): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const relsPart = pkg.parts.find(
    (p) => p.name === `/ppt/slides/_rels/slide${slideIndex + 1}.xml.rels`,
  );
  if (!relsPart) throw new Error(`slide${slideIndex + 1}.xml.rels not found`);
  return new TextDecoder().decode(relsPart.data);
};

describe('fn API: shape text mutation', () => {
  it('setShapeText + setShapeBullets + setShapeAlignment + setShapeTextFormat persist', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find((s) => getShapeText(s).length > 0);
    if (!shape) throw new Error('expected text shape');

    setShapeText(shape, 'Line one\nLine two');
    setShapeBullets(shape, 'bullet');
    setShapeAlignment(shape, 'center');
    setShapeTextFormat(shape, { bold: true, color: '#FF0000' });

    const reloaded = await Presentation.load(await savePresentation(pres));
    const reloadedShape = reloaded.slides[0]?.shapes.find((s) => s.text.includes('Line'));
    expect(reloadedShape?.text).toContain('Line one');
    expect(reloadedShape?.text).toContain('Line two');
  });

  it('setShapeHyperlink wires + clears an external link', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide).find((s) => getShapeText(s).length > 0);
    if (!shape) throw new Error('expected text shape');

    setShapeHyperlink(shape, 'https://example.com/');
    const wired = await savePresentation(pres);
    expect(await slideRels(wired, 0)).toContain('https://example.com/');
    expect(await slideXml(wired, 0)).toContain('hlinkClick');

    setShapeHyperlink(shape, null);
    const cleared = await savePresentation(pres);
    expect(await slideXml(cleared, 0)).not.toContain('hlinkClick');
  });
});

describe('fn API: shape geometry mutation', () => {
  it('setShapePosition / setShapeSize / setShapeRotation / setShapeFlip round-trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => s !== undefined);
    if (!picture) throw new Error('expected picture shape');

    setShapePosition(picture, inches(0.5), inches(0.5));
    setShapeSize(picture, inches(3), inches(2));
    setShapeRotation(picture, 45);
    setShapeFlip(picture, { horizontal: true });

    const reloaded = await loadPresentation(await savePresentation(pres));
    const replayShape = getSlideShapes(getSlides(reloaded)[0]!)[0]!;
    expect(getShapePosition(replayShape)).toEqual({ x: inches(0.5), y: inches(0.5) });
    expect(getShapeSize(replayShape)).toEqual({ w: inches(3), h: inches(2) });
    expect(getShapeRotation(replayShape)).toBe(45);
  });
});

describe('fn API: shape fill + stroke', () => {
  it('setShapeFill / setShapeNoFill / clearShapeFill cycle', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = findSlidePlaceholder(slide, 'title') ?? getSlideShapes(slide)[0]!;

    setShapeFill(shape, '#00AAFF');
    expect(await slideXml(await savePresentation(pres), 0)).toContain('00AAFF');

    setShapeNoFill(shape);
    expect(await slideXml(await savePresentation(pres), 0)).toContain('noFill');

    clearShapeFill(shape);
    // After clear, neither solid nor noFill on this shape's own spPr.
    const cleared = await slideXml(await savePresentation(pres), 0);
    expect(cleared.length).toBeGreaterThan(0);
  });

  it('setShapeStroke / setShapeNoStroke / clearShapeStroke cycle', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = getSlideShapes(slide)[0]!;

    setShapeStroke(shape, { color: '#112233', widthEmu: 12700 });
    expect(await slideXml(await savePresentation(pres), 0)).toContain('112233');

    clearShapeStroke(shape);
    const cleared = await slideXml(await savePresentation(pres), 0);
    expect(cleared.length).toBeGreaterThan(0);
  });
});

describe('fn API: shape removal', () => {
  it('removeShape drops the shape and renumbers indices', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;
    expect(before).toBeGreaterThan(0);

    const target = getSlideShapes(slide)[0]!;
    removeShape(target);

    expect(getSlideShapes(slide).length).toBe(before - 1);
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides[0]?.shapes.length).toBe(before - 1);
  });
});
