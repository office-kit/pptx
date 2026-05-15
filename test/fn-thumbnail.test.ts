// getThumbnail / setThumbnail / removeThumbnail — package-level
// thumbnail image (the one PowerPoint, Finder, etc. show as the
// file preview).

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  getThumbnail,
  listPackageParts,
  loadPresentation,
  removeThumbnail,
  savePresentation,
  setThumbnail,
  type PresentationData,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const hasPart = (pres: PresentationData, partName: string): boolean =>
  listPackageParts(pres).some((p) => p.name === partName);

describe('fn API: thumbnail helpers', () => {
  it('reads the JPEG thumbnail shipped with the fixture', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const thumb = getThumbnail(pres);
    expect(thumb).not.toBeNull();
    expect(thumb!.format).toBe('jpeg');
    expect(thumb!.bytes.byteLength).toBeGreaterThan(0);
  });

  it('replaces the thumbnail in place when the format matches', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    // Replace JPEG → PNG; the part name should change to .png.
    setThumbnail(pres, PNG);
    const thumb = getThumbnail(pres);
    expect(thumb).not.toBeNull();
    expect(thumb!.format).toBe('png');

    const reloaded = await loadPresentation(await savePresentation(pres));
    const thumb2 = getThumbnail(reloaded);
    expect(thumb2!.format).toBe('png');
    expect(thumb2!.bytes.byteLength).toBe(PNG.byteLength);
  });

  it('bootstraps a thumbnail when the package has none', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    removeThumbnail(pres);
    expect(getThumbnail(pres)).toBeNull();

    setThumbnail(pres, PNG);
    const thumb = getThumbnail(pres);
    expect(thumb).not.toBeNull();
    expect(thumb!.format).toBe('png');

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(hasPart(reloaded, '/docProps/thumbnail.png')).toBe(true);
  });

  it('removeThumbnail clears the part and the rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    expect(getThumbnail(pres)).not.toBeNull();
    removeThumbnail(pres);
    expect(getThumbnail(pres)).toBeNull();

    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getThumbnail(reloaded)).toBeNull();
    expect(hasPart(reloaded, '/docProps/thumbnail.jpeg')).toBe(false);
  });
});
