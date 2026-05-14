// Free-function slide-level authoring API: addSlideTextBox / addSlideShape /
// addSlideLine / addSlideTable / addSlideImage, plus background, transition,
// notes, and setShapeImage.

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import {
  Presentation,
  _internalPackageOf,
  addSlideImage,
  addSlideLine,
  addSlideShape,
  addSlideTable,
  addSlideTextBox,
  clearSlideBackground,
  clearSlideTransition,
  getShapeKind,
  getShapeText,
  getSlideNotes,
  getSlideShapes,
  getSlides,
  inches,
  loadPresentation,
  savePresentation,
  setShapeImage,
  setSlideBackground,
  setSlideNotes,
  setSlideTransition,
} from '../src/api/index.ts';

const fixture = (name: string): string =>
  fileURLToPath(new URL(`./fixtures/minimal/${name}`, import.meta.url));

const slideXml = async (bytes: Uint8Array, slideIndex: number): Promise<string> => {
  const pres = await Presentation.load(bytes);
  const pkg = _internalPackageOf(pres);
  const slidePart = pkg.parts.find(
    (p) => p.name === `/ppt/slides/slide${slideIndex + 1}.xml`,
  );
  if (!slidePart) throw new Error(`slide${slideIndex + 1}.xml not found`);
  return new TextDecoder().decode(slidePart.data);
};

const tinyPng = (): Uint8Array =>
  new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
  ]);

const tinyJpeg = (): Uint8Array =>
  new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0xff, 0xd9]);

describe('fn API: slide authoring', () => {
  it('addSlideTextBox appends a new text-bearing shape', async () => {
    const pres = await loadPresentation(await readFile(fixture('blank.pptx')));
    // blank.pptx has no slides; add one via the class API to get a target slide.
    const seed = await Presentation.load(await savePresentation(pres));
    const layout = seed.slideLayouts.find((l) => l.name === 'Title Only');
    if (!layout) throw new Error('Title Only layout missing');
    seed.addSlide({ layout });
    const reloaded = await loadPresentation(await seed.save());
    const slide = getSlides(reloaded)[0]!;
    const beforeCount = getSlideShapes(slide).length;

    const tb = addSlideTextBox(slide, {
      x: inches(1),
      y: inches(1),
      w: inches(3),
      h: inches(1),
      text: 'Hello textbox',
    });
    expect(getShapeKind(tb)).toBe('shape');
    expect(getShapeText(tb)).toBe('Hello textbox');
    expect(getSlideShapes(slide).length).toBe(beforeCount + 1);
  });

  it('addSlideShape adds a preset shape with text', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;

    const sp = addSlideShape(slide, {
      preset: 'ellipse',
      x: inches(0.5),
      y: inches(0.5),
      w: inches(2),
      h: inches(1),
      text: 'Ellipse',
    });
    expect(getShapeText(sp)).toBe('Ellipse');
    expect(getSlideShapes(slide).length).toBe(before + 1);
  });

  it('addSlideLine adds a connector', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;

    const ln = addSlideLine(slide, {
      from: { x: inches(0), y: inches(0) },
      to: { x: inches(3), y: inches(2) },
      color: '#FF0000',
    });
    expect(getShapeKind(ln)).toBe('connector');
    expect(getSlideShapes(slide).length).toBe(before + 1);
  });

  it('addSlideTable adds a graphic-frame table', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;

    const tbl = addSlideTable(slide, {
      x: inches(0),
      y: inches(0),
      w: inches(4),
      h: inches(2),
      rows: [
        ['A', 'B'],
        ['C', 'D'],
      ],
    });
    expect(getShapeKind(tbl)).toBe('graphicFrame');
    expect(getSlideShapes(slide).length).toBe(before + 1);
  });

  it('addSlideImage adds a picture and registers the media part', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    const before = getSlideShapes(slide).length;

    const pic = addSlideImage(slide, tinyPng(), {
      x: inches(0),
      y: inches(0),
      w: inches(1),
      h: inches(1),
    });
    expect(getShapeKind(pic)).toBe('picture');
    expect(getSlideShapes(slide).length).toBe(before + 1);

    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides[0]?.shapes.length).toBe(before + 1);
  });
});

describe('fn API: slide background + transition', () => {
  it('setSlideBackground + clearSlideBackground cycle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    setSlideBackground(slide, '#112233');
    expect(await slideXml(await savePresentation(pres), 0)).toContain('112233');

    clearSlideBackground(slide);
    const cleared = await slideXml(await savePresentation(pres), 0);
    expect(cleared).not.toContain('<p:bg>');
  });

  it('setSlideTransition + clearSlideTransition cycle', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;

    setSlideTransition(slide, { effect: 'fade', speed: 'med' });
    const wired = await slideXml(await savePresentation(pres), 0);
    expect(wired).toContain('<p:transition');
    expect(wired).toContain('<p:fade');

    clearSlideTransition(slide);
    const cleared = await slideXml(await savePresentation(pres), 0);
    expect(cleared).not.toContain('<p:transition');
  });
});

describe('fn API: slide notes', () => {
  it('setSlideNotes + getSlideNotes round-trip', async () => {
    const pres = await loadPresentation(await readFile(fixture('two-slides.pptx')));
    const slide = getSlides(pres)[0]!;
    expect(getSlideNotes(slide)).toBeNull();

    setSlideNotes(slide, 'Speaker notes line one\nLine two');
    const reloaded = await loadPresentation(await savePresentation(pres));
    expect(getSlideNotes(getSlides(reloaded)[0]!)).toBe('Speaker notes line one\nLine two');
  });
});

describe('fn API: shape image replacement', () => {
  it('setShapeImage in-place same-format', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected picture shape');

    setShapeImage(picture, tinyPng(), { format: 'png' });
    const reloaded = await Presentation.load(await savePresentation(pres));
    const reloadedPic = reloaded.slides[0]?.shapes.find((s) => s.kind === 'picture');
    expect(reloadedPic).toBeDefined();
  });

  it('setShapeImage cross-format allocates new media + repoints rel', async () => {
    const pres = await loadPresentation(await readFile(fixture('one-image-slide.pptx')));
    const slide = getSlides(pres)[0]!;
    const picture = getSlideShapes(slide).find((s) => getShapeKind(s) === 'picture');
    if (!picture) throw new Error('expected picture shape');

    setShapeImage(picture, tinyJpeg(), { format: 'jpeg' });
    const reloaded = await Presentation.load(await savePresentation(pres));
    expect(reloaded.slides[0]?.shapes.length).toBeGreaterThan(0);
  });
});
