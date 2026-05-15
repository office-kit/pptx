// setMediaPartBytes — bulk-swap an image used everywhere.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getMediaParts,
  getSlides,
  inches,
  loadPresentation,
  setMediaPartBytes,
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

const otherPng = (): Uint8Array => {
  // Same shape as tinyPng but a distinct byte to verify the swap.
  const b = tinyPng();
  b[10] = (b[10] ?? 0) ^ 0xff;
  return b;
};

describe('fn API: setMediaPartBytes', () => {
  it('replaces bytes for an existing media part', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), format: 'png',
    });
    const media = getMediaParts(pres);
    expect(media.length).toBeGreaterThan(0);
    const target = media[0]!;
    const ok = setMediaPartBytes(pres, target.name, otherPng());
    expect(ok).toBe(true);
    expect(getMediaParts(pres)[0]!.data).toEqual(otherPng());
  });

  it('returns false when the part name is unknown', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(setMediaPartBytes(pres, '/ppt/media/missing.png', tinyPng())).toBe(false);
  });
});
