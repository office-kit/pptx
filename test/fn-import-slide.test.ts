// importSlide — copy a slide from one deck into another.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  _internalPackageOf,
  addSlideChart,
  getShapeChartSpec,
  setSlideNotes,
  getSlideNotes,
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

  it('preserves charts, workbooks and Japanese/English speaker notes across save/load', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(source)[0]!;
    addSlideChart(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(5),
      h: inches(3),
      spec: {
        kind: 'column',
        categories: ['日本語', 'English'],
        series: [{ name: 'Sales', values: [3, 7] }],
      },
    });
    setSlideNotes(slide, '発表メモ\nSpeaker notes');
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const imported = importSlide(target, slide, findSlideLayout(target, 'Blank')!);
    expect(getSlideNotes(imported)).toBe('発表メモ\nSpeaker notes');
    const reloaded = await loadPresentation(await savePresentation(target));
    const result = getSlides(reloaded)[0]!;
    const chart = getSlideShapes(result).find((shape) => getShapeChartSpec(shape) !== null)!;
    expect(getShapeChartSpec(chart)?.categories).toEqual(['日本語', 'English']);
    expect(getShapeChartSpec(chart)?.series[0]?.values).toEqual([3, 7]);
    expect(getSlideNotes(result)).toBe('発表メモ\nSpeaker notes');
    expect(
      _internalPackageOf(reloaded).parts.some((part) =>
        part.contentType.includes('spreadsheetml.sheet'),
      ),
    ).toBe(true);
  });

  it('keeps repeated imports independent, including notes and their slide back-references', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const sourceSlide = getSlides(source)[0]!;
    setSlideNotes(sourceSlide, 'Original notes');
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const layout = findSlideLayout(target, 'Blank')!;
    const first = importSlide(target, sourceSlide, layout);
    const second = importSlide(target, sourceSlide, layout);
    setSlideNotes(first, 'Changed notes');
    expect(getSlideNotes(second)).toBe('Original notes');
    expect(getSlideNotes(sourceSlide)).toBe('Original notes');
    const reloaded = await loadPresentation(await savePresentation(target));
    expect(getSlides(reloaded).map(getSlideNotes)).toEqual(['Changed notes', 'Original notes']);
    const pkg = _internalPackageOf(reloaded);
    const notes = pkg.parts.filter((part) => part.contentType.endsWith('notesSlide+xml'));
    expect(notes).toHaveLength(2);
    const destinations = notes.map(
      (part) => pkg.getRels(part.name)!.items.find((rel) => rel.type.endsWith('/slide'))!.target,
    );
    expect(new Set(destinations)).toEqual(
      new Set(['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml']),
    );
  });

  it('does not change the target when an imported dependency is missing', async () => {
    const source = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const sourceSlide = getSlides(source)[0]!;
    addSlideImage(sourceSlide, tinyPng(), {
      x: inches(1),
      y: inches(1),
      w: inches(1),
      h: inches(1),
    });
    const sourcePkg = _internalPackageOf(source);
    sourcePkg.removePart(partName(getMediaParts(source)[0]!.name));
    const target = await loadPresentation(await readFile(fixture('blank.pptx')));
    const before = _internalPackageOf(target).parts.map((part) => ({
      ...part,
      data: new Uint8Array(part.data),
    }));
    expect(() => importSlide(target, sourceSlide, findSlideLayout(target, 'Blank')!)).toThrow(
      /missing dependency/,
    );
    expect(_internalPackageOf(target).parts).toEqual(before);
    expect(getSlides(target)).toHaveLength(0);
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
