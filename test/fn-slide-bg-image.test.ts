// setSlideBackgroundImage — picture as slide background.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  clearSlideBackground,
  getMediaParts,
  getSlideBackground,
  getSlides,
  loadPresentation,
  savePresentation,
  setSlideBackgroundImage,
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

describe('fn API: setSlideBackgroundImage', () => {
  it('writes a blipFill background and allocates a media part', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackgroundImage(slide, tinyPng(), { format: 'png' });

    expect(getSlideBackground(slide).kind).toBe('image');

    const bytes = await savePresentation(pres);
    const reloaded = await loadPresentation(bytes);
    expect(getMediaParts(reloaded).some((p) => /^\/ppt\/media\/image\d+\.png$/.test(p.name))).toBe(
      true,
    );
  });

  it('clearSlideBackground removes the picture background', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    setSlideBackgroundImage(slide, tinyPng(), { format: 'png' });
    clearSlideBackground(slide);
    expect(getSlideBackground(slide).kind).toBe('inherit');
  });
});
