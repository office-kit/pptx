// copyShape — clone a shape onto another slide in the same deck.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  addSlideShape,
  copyShape,
  findShapesByKind,
  getShapeKind,
  getShapeText,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

describe('fn API: copyShape', () => {
  it('copies a text shape onto another slide', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const sp = addSlideShape(slideA!, {
      preset: 'rect',
      x: inches(1),
      y: inches(1),
      w: inches(2),
      h: inches(1),
      text: 'COPY ME',
      name: 'Tagged',
    });
    const beforeBCount = getSlideShapes(slideB!).length;

    const copied = copyShape(slideB!, sp);
    expect(getShapeKind(copied)).toBe('shape');
    expect(getShapeText(copied)).toBe('COPY ME');
    expect(getSlideShapes(slideB!).length).toBe(beforeBCount + 1);
  });

  it('copies a picture shape and the new shape still references valid media', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const [slideA, slideB] = getSlides(pres);
    const pic = addSlideImage(slideA!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
      format: 'png',
    });
    const copied = copyShape(slideB!, pic);
    expect(getShapeKind(copied)).toBe('picture');
    // The target slide now has at least one picture.
    expect(findShapesByKind(slideB!, 'picture').length).toBeGreaterThan(0);
  });

  it('rejects cross-package copy', async () => {
    const presA = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const presB = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const sourceShape = getSlideShapes(getSlides(presA)[0]!)[0]!;
    const targetSlide = getSlides(presB)[0]!;
    expect(() => copyShape(targetSlide, sourceShape)).toThrow(/same package/);
  });
});
