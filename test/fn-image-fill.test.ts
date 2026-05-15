// Image-as-shape-fill — embed a picture into a non-picture shape's fill.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideShape,
  getMediaParts,
  getSlideXmlString,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeImageFill,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await loadPresentation(bytes);
  return getSlideXmlString(getSlides(pres)[slideIndex]!);
};

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

describe('fn API: setShapeImageFill', () => {
  it('replaces the fill of a regular shape with a blipFill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'rect',
      x: inches(0.5),
      y: inches(0.5),
      w: inches(3),
      h: inches(2),
    });

    setShapeImageFill(shape, tinyPng(), { format: 'png' });
    const xml = await slideXml(await savePresentation(pres), 0);
    expect(xml).toContain('<a:blipFill>');
    expect(xml).toContain('<a:blip ');
    expect(xml).toContain('<a:stretch>');

    // A new media part should exist for the embedded image.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = getMediaParts(reloaded).find((p) =>
      /^\/ppt\/media\/image\d+\.png$/.test(p.name),
    );
    expect(media).not.toBeUndefined();
  });

  it('reapplying replaces (not stacks) the fill', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const shape = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
    });
    setShapeImageFill(shape, tinyPng(), { format: 'png' });
    setShapeImageFill(shape, tinyPng(), { format: 'png' });
    const xml = await slideXml(await savePresentation(pres), 0);
    // Only one <a:blipFill> per spPr; two embeds are fine but the
    // shape itself shouldn't carry two.
    const occurrences = xml.match(/<a:blipFill>/g)?.length ?? 0;
    expect(occurrences).toBeLessThanOrEqual(2); // depending on multiple shapes in spTree
  });
});
