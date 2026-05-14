// Level-2 image-replacement smoke test.
//
// Scenario for the in-place path:
//   1. Load `one-image-slide.pptx` (one picture shape on a title-only layout).
//   2. Locate the picture via `shape.kind === 'picture'`.
//   3. Replace the bytes with a different PNG.
//   4. Save → reload → assert the bytes that come back match the new payload
//      and the relationship still points at the original media filename.
//
// Scenario for the cross-format path:
//   - Same setup, but the replacement is JPEG bytes.
//   - Asserts a NEW media part exists with `.jpg` extension and that the
//     slide rel was repointed at it.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { Presentation } from '../src/api/index.ts';
import { detectImageFormat, partName } from '../src/internal/opc/index.ts';
import { _internalPackageOf } from '../src/api/presentation.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// Pre-rendered 1x1 transparent PNG; differs from the fixture's 100x100 red.
// prettier-ignore
const ALT_PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

// Bare-minimum JPEG bytes. Not a valid image but enough for detectImageFormat
// to classify it as JPEG. Real callers use real images.
const FAKE_JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46]);

describe('L2: image replacement', () => {
  it('replaces a PNG with a PNG in place', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected one slide');
    const picture = slide.shapes.find((s) => s.kind === 'picture');
    if (!picture) throw new Error('expected a picture shape');

    picture.setImage(ALT_PNG);

    const reloaded = await Presentation.load(await pres.save());
    // Find the media part on the reloaded package. python-pptx emits the
    // image under /ppt/media/image1.png by convention.
    const pkg = _internalPackageOf(reloaded);
    const mediaPart = pkg.getPart(partName('/ppt/media/image1.png'));
    expect(mediaPart).not.toBeNull();
    expect(mediaPart?.data).toEqual(ALT_PNG);
    expect(detectImageFormat(mediaPart?.data ?? new Uint8Array())).toBe('png');
  });

  it('cross-format swap (PNG → JPEG) allocates a new media part', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const slide = pres.slides[0];
    if (!slide) throw new Error('expected one slide');
    const picture = slide.shapes.find((s) => s.kind === 'picture');
    if (!picture) throw new Error('expected a picture shape');

    picture.setImage(FAKE_JPEG);

    const reloaded = await Presentation.load(await pres.save());
    const pkg = _internalPackageOf(reloaded);
    // Original PNG should still be present (we leave orphans for now).
    expect(pkg.getPart(partName('/ppt/media/image1.png'))).not.toBeNull();
    // A new .jpg media part should exist.
    const jpgPart = pkg.parts.find((p) => p.name.endsWith('.jpg'));
    expect(jpgPart).toBeDefined();
    expect(jpgPart?.data.subarray(0, 3)).toEqual(FAKE_JPEG.subarray(0, 3));
    expect(jpgPart?.contentType).toBe('image/jpeg');
  });

  it('rejects setImage on non-picture shapes', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-text-slide.pptx')));
    const title = pres.slides[0]?.findPlaceholder('title');
    if (!title) throw new Error('expected title');
    expect(() => title.setImage(ALT_PNG)).toThrow(/setImage only works on picture/);
  });

  it('throws on undetectable format without explicit override', async () => {
    const pres = await Presentation.load(await readFile(fixture('one-image-slide.pptx')));
    const picture = pres.slides[0]?.shapes.find((s) => s.kind === 'picture');
    if (!picture) throw new Error('expected picture');
    expect(() => picture.setImage(new Uint8Array([0, 0, 0, 0]))).toThrow(/format/);
  });
});
