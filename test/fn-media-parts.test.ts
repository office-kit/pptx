// getMediaParts — enumerate every /ppt/media/... part.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getMediaParts,
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

describe('fn API: getMediaParts', () => {
  it('returns embedded picture parts after addSlideImage', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideImage(slide, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
      format: 'png',
    });
    const media = getMediaParts(pres);
    expect(media.length).toBeGreaterThan(0);
    expect(media.every((m) => m.name.startsWith('/ppt/media/'))).toBe(true);
    expect(media.some((m) => m.contentType.includes('png'))).toBe(true);
    // The bytes should round-trip.
    const found = media.find((m) => m.name.endsWith('.png'));
    expect(found?.data.byteLength).toBe(tinyPng().byteLength);
  });

  it('returns empty for a deck without media', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // two-slides.pptx has no media — verify.
    expect(getMediaParts(pres)).toEqual([]);
  });
});
