// importSlide — copy a slide from one deck into another.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlide,
  addSlideImage,
  findSlideLayout,
  getMediaParts,
  getSlideLayout,
  getSlideLayoutName,
  getSlideShapes,
  getSlideText,
  getSlides,
  importSlide,
  inches,
  loadPresentation,
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

describe('fn API: importSlide', () => {
  it('copies a slide from one deck to another', async () => {
    // Source deck.
    const source = await loadPresentation(await readFile(fixture('one-text-slide.pptx')));
    const sourceText = getSlideText(getSlides(source)[0]!);

    // Target deck.
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Title and Content');
    expect(layout).not.toBeNull();

    const imported = importSlide(target, getSlides(source)[0]!, layout!);
    expect(imported).toBeDefined();
    expect(getSlideText(imported)).toBe(sourceText);
    expect(getSlides(target)).toHaveLength(1);

    const reloaded = await loadPresentation(await savePresentation(target));
    expect(getSlides(reloaded).length).toBe(1);
  });

  it('copies image media along with the slide', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(source)[0]!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
      format: 'png',
    });

    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Title and Content')!;
    expect(getMediaParts(target)).toEqual([]);

    importSlide(target, getSlides(source)[0]!, layout);

    const media = getMediaParts(target);
    expect(media.length).toBeGreaterThan(0);
    expect(media.some((m) => m.contentType.includes('png'))).toBe(true);
  });

  it('gives the layout rel a unique id when the source layout rel is not rId1', async () => {
    // Real templates often store the layout rel at rId3 with an image at rId1.
    // The new slide's layout rel must not collide with a preserved image rel id —
    // a hardcoded "rId1" would emit two relationships with the same Id.
    const src0 = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    addSlideImage(getSlides(src0)[0]!, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(2),
      h: inches(2),
      format: 'png',
    });
    const source = await loadPresentation(await savePresentation(src0));
    const srcPkg = _internalPackageOf(source);
    const srcSlide = partName('/ppt/slides/slide1.xml');
    const srcRels = srcPkg.getRels(srcSlide)!;
    // Move the image rel onto rId1 and the layout rel onto rId3, recreating the
    // collision the old hardcoded-rId1 code produced.
    srcPkg.setRels(srcSlide, {
      items: srcRels.items.map((r) =>
        r.target.includes('/media/')
          ? { ...r, id: 'rId1' }
          : r.target.includes('slideLayout')
            ? { ...r, id: 'rId3' }
            : r,
      ),
    });

    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Title and Content')!;
    importSlide(target, getSlides(source)[0]!, layout);

    // Target started blank, so the only slide rels part is the imported slide.
    const relsPart = _internalPackageOf(target).parts.find((p) =>
      /\/ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(p.name),
    )!;
    const ids = [...new TextDecoder().decode(relsPart.data).matchAll(/Id="([^"]+)"/g)].map(
      (m) => m[1],
    );
    expect(ids.length).toBeGreaterThanOrEqual(2);
    expect(new Set(ids).size).toBe(ids.length);
  });

  it('binds the imported slide to the supplied layout', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Blank')!;
    addSlide(target, { layout: findSlideLayout(target, 'Title Slide')! });

    importSlide(target, getSlides(source)[0]!, layout);

    const reloaded = await loadPresentation(await savePresentation(target));
    const last = getSlides(reloaded).at(-1)!;
    const lastLayout = getSlideLayout(last);
    expect(lastLayout && getSlideLayoutName(lastLayout)).toBe('Blank');
    expect(getSlideShapes(getSlides(target).at(-1)!).length).toBeGreaterThan(0);
  });
});
