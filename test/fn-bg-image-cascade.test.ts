// Background image cascade readers — verify that
// `getSlideBackgroundImageBytes` / `getSlideLayoutBackgroundImageBytes` /
// `getSlideMasterBackgroundImageBytes` return `null` for slides /
// layouts / masters that don't author a `<p:bgPr><a:blipFill>`, and
// that the slide-level reader returns bytes after
// `setSlideBackgroundImage`.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getSlideBackgroundImageBytes,
  getSlideLayout,
  getSlideLayoutBackgroundImageBytes,
  getSlideMasterBackgroundImageBytes,
  getSlides,
  loadPresentation,
  setSlideBackgroundImage,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const pngBytes = (): Uint8Array =>
  // 1x1 red PNG.
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0xa1, 0x5e, 0x10, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
  ]);

describe('fn API: background image cascade', () => {
  it('returns null on all three levels for slides without authored bg images', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    for (const slide of getSlides(pres)) {
      const layout = getSlideLayout(slide);
      expect(getSlideBackgroundImageBytes(slide)).toBeNull();
      if (layout) {
        expect(getSlideLayoutBackgroundImageBytes(pres, layout)).toBeNull();
        expect(getSlideMasterBackgroundImageBytes(pres, layout)).toBeNull();
      }
    }
  });

  it('slide reader returns bytes after setSlideBackgroundImage', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideBackgroundImageBytes(slide)).toBeNull();
    setSlideBackgroundImage(slide, pngBytes(), { format: 'png' });
    const bytes = getSlideBackgroundImageBytes(slide);
    expect(bytes).not.toBeNull();
    expect(bytes!.length).toBeGreaterThan(0);
  });
});
