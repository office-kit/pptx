// Level-2 image-replacement smoke test.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  findSlidePlaceholder,
  getMediaParts,
  getShapeKind,
  getSlideShapes,
  getSlides,
  loadPresentation,
  savePresentation,
  setShapeImage,
} from '../src/api/index.ts';
import { detectImageFormat } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const ALT_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const FAKE_JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46]);

describe('L2: image replacement', () => {
  it('replaces a PNG with a PNG in place', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected a picture shape');

    setShapeImage(picture, ALT_PNG);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = getMediaParts(reloaded).find((p) => p.name === '/ppt/media/image1.png');
    expect(media).toBeDefined();
    expect(media?.data).toEqual(ALT_PNG);
    expect(detectImageFormat(media?.data ?? new Uint8Array())).toBe('png');
  });

  it('cross-format swap (PNG → JPEG) allocates a new media part', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected a picture shape');

    setShapeImage(picture, FAKE_JPEG);

    const reloaded = await loadPresentation(await savePresentation(pres));
    const media = getMediaParts(reloaded);
    expect(media.find((p) => p.name === '/ppt/media/image1.png')).toBeDefined();
    const jpg = media.find((p) => p.name.endsWith('.jpg'));
    expect(jpg).toBeDefined();
    expect(jpg?.data.subarray(0, 3)).toEqual(FAKE_JPEG.subarray(0, 3));
    expect(jpg?.contentType).toBe('image/jpeg');
  });

  it('rejects setShapeImage on non-picture shapes', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const title = findSlidePlaceholder(getSlides(pres)[0]!, 'title');
    if (!title) throw new Error('expected title');
    expect(() => setShapeImage(title, ALT_PNG)).toThrow(/setShapeImage only works on picture/);
  });

  it('throws on undetectable format without explicit override', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const picture = getSlideShapes(getSlides(pres)[0]!).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected picture');
    expect(() => setShapeImage(picture, new Uint8Array([0, 0, 0, 0]))).toThrow(/format/);
  });
});
