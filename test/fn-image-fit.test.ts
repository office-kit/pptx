// addSlideImage / setShapeImage `fit` option — 'contain' scales the image
// to fit inside the target box preserving aspect ratio (centered), 'fill'
// (the default) stretches to the exact box as before. Natural size comes
// from the PNG / JPEG header; unmeasurable formats fall back to 'fill'.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  addSlideImage,
  getShapeBounds,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeImage,
} from '../src/api/index.ts';
import { readImagePixelSize } from '../src/internal/opc/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

// prettier-ignore
const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

// The size reader only parses the IHDR header, so a dimension-patched copy
// of the 1×1 PNG is enough to exercise the fit math (the pixel payload is
// stored opaquely; nothing in the save path re-decodes it).
const pngWithSize = (width: number, height: number): Uint8Array => {
  const bytes = PNG_1X1.slice();
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
};

// Minimal JPEG: SOI + one SOF0 segment carrying 64×32 + EOI.
// prettier-ignore
const JPEG_64X32 = new Uint8Array([
  0xff, 0xd8,
  0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x20, 0x00, 0x40, 0x03,
  0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
  0xff, 0xd9,
]);

// prettier-ignore
const GIF_HEADER = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x02, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00,
  0x00, 0x00, 0x00, 0xff, 0xff, 0xff, 0x3b,
]);

describe('internal: readImagePixelSize', () => {
  it('reads PNG dimensions from the IHDR header', () => {
    expect(readImagePixelSize(pngWithSize(200, 100))).toEqual({ width: 200, height: 100 });
    expect(readImagePixelSize(PNG_1X1)).toEqual({ width: 1, height: 1 });
  });

  it('reads JPEG dimensions from the SOF marker', () => {
    expect(readImagePixelSize(JPEG_64X32)).toEqual({ width: 64, height: 32 });
  });

  it('returns null for formats without a cheap header (GIF) and truncated bytes', () => {
    expect(readImagePixelSize(GIF_HEADER)).toBeNull();
    expect(readImagePixelSize(PNG_1X1.subarray(0, 20))).toBeNull();
    expect(readImagePixelSize(new Uint8Array([0xff, 0xd8, 0xff, 0xd9]))).toBeNull();
  });
});

describe('fn API: addSlideImage fit option', () => {
  const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(2) };

  it("defaults to 'fill' — the picture takes the exact box", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, pngWithSize(200, 100), box);
    expect(getShapeBounds(pic)).toEqual({ x: box.x, y: box.y, w: box.w, h: box.h });
  });

  it("'contain' letterboxes a wide image (full width, centered vertically)", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, pngWithSize(200, 100), { ...box, fit: 'contain' });
    const b = getShapeBounds(pic)!;
    expect(b.w).toBe(inches(2));
    expect(b.h).toBe(inches(1));
    expect(b.x).toBe(inches(1));
    // Centered: y + (boxH - h) / 2 = 1in + 0.5in.
    expect(b.y).toBe(inches(1.5));
  });

  it("'contain' pillarboxes a tall image (full height, centered horizontally)", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, pngWithSize(100, 200), { ...box, fit: 'contain' });
    const b = getShapeBounds(pic)!;
    expect(b.w).toBe(inches(1));
    expect(b.h).toBe(inches(2));
    expect(b.x).toBe(inches(1.5));
    expect(b.y).toBe(inches(1));
  });

  it("'contain' honors a JPEG's SOF dimensions", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, JPEG_64X32, { ...box, fit: 'contain' });
    const b = getShapeBounds(pic)!;
    expect(b.w).toBe(inches(2));
    expect(b.h).toBe(inches(1));
  });

  it("'contain' falls back to 'fill' for unmeasurable formats (GIF) without erroring", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, GIF_HEADER, { ...box, fit: 'contain' });
    expect(getShapeBounds(pic)).toEqual({ x: box.x, y: box.y, w: box.w, h: box.h });
  });

  it("'contain' geometry survives save → load", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    addSlideImage(slide, pngWithSize(200, 100), { ...box, fit: 'contain' });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pics = getSlides(reloaded)[0]!;
    const { getSlideShapes } = await import('../src/api/index.ts');
    const pic = getSlideShapes(pics).find((s) => getShapeBounds(s)?.h === inches(1));
    expect(pic).not.toBeUndefined();
  });
});

describe('fn API: setShapeImage fit option', () => {
  const box = { x: inches(1), y: inches(1), w: inches(2), h: inches(2) };

  it('preserves geometry by default (back-compat)', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG_1X1, box);
    setShapeImage(pic, pngWithSize(200, 100));
    expect(getShapeBounds(pic)).toEqual({ x: box.x, y: box.y, w: box.w, h: box.h });
  });

  it("'contain' re-fits the extent to the replacement image inside the current box", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG_1X1, box);
    setShapeImage(pic, pngWithSize(200, 100), { fit: 'contain' });
    const b = getShapeBounds(pic)!;
    expect(b.w).toBe(inches(2));
    expect(b.h).toBe(inches(1));
    expect(b.y).toBe(inches(1.5));
  });

  it("'contain' with an unmeasurable replacement leaves the box unchanged", async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, PNG_1X1, box);
    setShapeImage(pic, GIF_HEADER, { fit: 'contain' });
    expect(getShapeBounds(pic)).toEqual({ x: box.x, y: box.y, w: box.w, h: box.h });
  });
});
