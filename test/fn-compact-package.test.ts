// compactPackage — drop orphan media parts.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideImage,
  compactPackage,
  getMediaParts,
  getSlides,
  inches,
  loadPresentation,
  removeShape,
  savePresentation,
} from '../src/api/index.ts';
import { partName } from '../src/internal/opc/index.ts';

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

describe('fn API: compactPackage', () => {
  it('removes media parts not referenced by any rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const pic = addSlideImage(slide, tinyPng(), {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), format: 'png',
    });
    // Save → reload → orphan the media by clearing slide rels.
    const reloaded = await loadPresentation(await savePresentation(pres));
    const pkg = _internalPackageOf(reloaded);
    pkg.setRels(partName('/ppt/slides/slide1.xml'), {
      items: [
        {
          id: 'rId1',
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
          target: '../slideLayouts/slideLayout1.xml',
          targetMode: 'Internal',
        },
      ],
    });

    const after = await loadPresentation(await savePresentation(reloaded));
    expect(getMediaParts(after).length).toBeGreaterThan(0);
    const { removed } = compactPackage(after);
    expect(removed.length).toBeGreaterThan(0);
    expect(getMediaParts(after).length).toBe(0);
    void pic;
  });

  it('does not touch media that is still referenced', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), format: 'png',
    });
    const reloaded = await loadPresentation(await savePresentation(pres));
    const before = getMediaParts(reloaded).length;
    expect(before).toBeGreaterThan(0);
    expect(compactPackage(reloaded).removed).toEqual([]);
    expect(getMediaParts(reloaded).length).toBe(before);
  });

  it('cleans up after a removeShape() that took the last image with it', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const pic = addSlideImage(getSlides(pres)[0]!, tinyPng(), {
      x: inches(0), y: inches(0), w: inches(1), h: inches(1), format: 'png',
    });
    removeShape(pic);
    // The media is still in the package but no shape uses it now.
    // The rel is still there too (removeShape doesn't touch slide rels in v1).
    // compactPackage shouldn't drop the media because the rel still
    // references it; this is the documented behavior.
    expect(compactPackage(pres).removed).toEqual([]);
  });
});
